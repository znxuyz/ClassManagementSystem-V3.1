/**
 * Atomic score operations — uses LockService to prevent race conditions
 * with 30 concurrent users.
 */

// Atomically add/subtract points for one student.
function atomicModifyStudentScore(studentId, delta, reason) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(e) {
    _logError('atomicModifyStudentScore-lock', e);
    return { success: false, error: 'lock_timeout' };
  }
  try {
    var sSh = _ensureExactSheet('Students', H.Students);
    var existing = _readFresh('Students', H.Students);
    var student = null;
    for (var i = 0; i < existing.length; i++) {
      if (String(existing[i].id) === String(studentId)) { student = existing[i]; break; }
    }
    if (!student) return { success: false, error: 'student_not_found' };

    student.score = (Number(student.score) || 0) + Number(delta);
    _upsertNoLock(sSh, 'Students', H.Students, existing, [student], 'id');

    var hist = {
      id: _makeHistoryId(),
      type: 'individual',
      targetId: studentId,
      targetName: student.name,
      groupName: student.groupName || '',
      scoreChange: Number(delta),
      reason: reason || (delta > 0 ? '個人加分' : '個人扣分'),
      date: _nowLocale(),
      affectedStudents: [{ id: student.id, name: student.name, scoreChange: Number(delta) }]
    };
    var hSh = _ensureExactSheet('ScoreHistory', H.ScoreHistory);
    _appendDirectNoLock(hSh, 'ScoreHistory', H.ScoreHistory, [hist], JSON_FIELDS.ScoreHistory);

    return { success: true, newScore: student.score };
  } catch(e) {
    _logError('atomicModifyStudentScore', e.toString());
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

// Atomically add/subtract points for a group and all its members.
function atomicModifyGroupScore(groupId, delta, reason) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(e) {
    _logError('atomicModifyGroupScore-lock', e);
    return { success: false, error: 'lock_timeout' };
  }
  try {
    var gSh = _ensureExactSheet('Groups', H.Groups);
    var sSh = _ensureExactSheet('Students', H.Students);

    var existingGroups = _readFresh('Groups', H.Groups);
    var existingStudents = _readFresh('Students', H.Students);

    var group = null;
    for (var i = 0; i < existingGroups.length; i++) {
      if (String(existingGroups[i].id) === String(groupId)) { group = existingGroups[i]; break; }
    }
    if (!group) return { success: false, error: 'group_not_found' };

    group.score = (Number(group.score) || 0) + Number(delta);
    _upsertNoLock(gSh, 'Groups', H.Groups, existingGroups, [group], 'id');

    var affected = [];
    var membersToUpdate = [];
    existingStudents.forEach(function(s) {
      if (String(s.groupId) === String(groupId)) {
        s.score = (Number(s.score) || 0) + Number(delta);
        membersToUpdate.push(s);
        affected.push({ id: s.id, name: s.name, scoreChange: Number(delta) });
      }
    });
    if (membersToUpdate.length) {
      _upsertNoLock(sSh, 'Students', H.Students, existingStudents, membersToUpdate, 'id');
    }

    var hist = {
      id: _makeHistoryId(),
      type: 'group',
      targetId: groupId,
      targetName: group.name,
      groupName: group.name,
      scoreChange: Number(delta),
      reason: reason || (delta > 0 ? '小組加分' : '小組扣分'),
      date: _nowLocale(),
      affectedStudents: affected
    };
    var hSh = _ensureExactSheet('ScoreHistory', H.ScoreHistory);
    _appendDirectNoLock(hSh, 'ScoreHistory', H.ScoreHistory, [hist], JSON_FIELDS.ScoreHistory);

    return { success: true, newGroupScore: group.score, affectedCount: affected.length };
  } catch(e) {
    _logError('atomicModifyGroupScore', e.toString());
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

// Atomically submit a quiz answer: check for duplicates, calculate rank, and award points if winner.
function atomicSubmitQuizAnswer(quizId, studentId, studentName, groupName, userAnswer) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(e) {
    _logError('atomicSubmitQuizAnswer-lock', e);
    return { success: false, error: 'lock_timeout' };
  }
  try {
    var qaSh = _ensureExactSheet('QuizAnswers', H.QuizAnswers);
    var existingAnswers = _readFresh('QuizAnswers', H.QuizAnswers);

    var alreadyAnswered = existingAnswers.some(function(a) {
      return String(a.quizId) === String(quizId) && String(a.studentId) === String(studentId);
    });
    if (alreadyAnswered) return { success: false, error: 'already_answered' };

    var existingQuizzes = _readFresh('Quizzes', H.Quizzes, JSON_FIELDS.Quizzes);
    var quiz = null;
    for (var i = 0; i < existingQuizzes.length; i++) {
      if (String(existingQuizzes[i].id) === String(quizId)) { quiz = existingQuizzes[i]; break; }
    }
    if (!quiz) return { success: false, error: 'quiz_not_found' };

    var correct = quiz.correct || quiz.answer || '';
    var isCorrect = (String(userAnswer) === String(correct));
    var rank = 0;
    var scoreAwarded = 0;

    if (isCorrect) {
      var correctSoFar = existingAnswers.filter(function(a) {
        return String(a.quizId) === String(quizId) && a.isCorrect;
      });
      var candidateRank = correctSoFar.length + 1;
      var winners = Number(quiz.winners) || 0;
      if (candidateRank <= winners) {
        rank = candidateRank;
        var scoresMap = quiz.scores || {};
        scoreAwarded = Number(scoresMap[String(rank)]) || 0;
      }
    }

    var answerRecord = {
      id: _makeHistoryId(),
      quizId: quizId,
      studentId: studentId,
      studentName: studentName,
      rank: rank,
      scoreAwarded: scoreAwarded,
      answer: userAnswer,
      isCorrect: isCorrect,
      submitTime: _nowISO()
    };
    _appendDirectNoLock(qaSh, 'QuizAnswers', H.QuizAnswers, [answerRecord]);

    if (scoreAwarded > 0) {
      var sSh = _ensureExactSheet('Students', H.Students);
      var existingStudents = _readFresh('Students', H.Students);
      var student = null;
      for (var si = 0; si < existingStudents.length; si++) {
        if (String(existingStudents[si].id) === String(studentId)) { student = existingStudents[si]; break; }
      }
      if (student) {
        student.score = (Number(student.score) || 0) + scoreAwarded;
        _upsertNoLock(sSh, 'Students', H.Students, existingStudents, [student], 'id');

        var hist = {
          id: _makeHistoryId(),
          type: 'quiz',
          targetId: studentId,
          targetName: studentName,
          groupName: groupName || '',
          scoreChange: scoreAwarded,
          reason: '搶分活動第' + rank + '名：' + (quiz.title || ''),
          date: _nowLocale(),
          affectedStudents: [{ id: studentId, name: studentName, scoreChange: scoreAwarded }]
        };
        var hSh = _ensureExactSheet('ScoreHistory', H.ScoreHistory);
        _appendDirectNoLock(hSh, 'ScoreHistory', H.ScoreHistory, [hist], JSON_FIELDS.ScoreHistory);
      }
    }

    return { success: true, isCorrect: isCorrect, rank: rank, scoreAwarded: scoreAwarded };
  } catch(e) {
    _logError('atomicSubmitQuizAnswer', e.toString());
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

// Atomically approve an exchange request: check score, deduct points, update reward quantity.
function atomicApproveExchange(requestId) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(e) {
    _logError('atomicApproveExchange-lock', e);
    return { success: false, error: 'lock_timeout' };
  }
  try {
    var exSh = _ensureExactSheet('ExchangeHistory', H.ExchangeHistory);
    var existingRequests = _readFresh('ExchangeHistory', H.ExchangeHistory);
    var request = null;
    for (var i = 0; i < existingRequests.length; i++) {
      if (String(existingRequests[i].id) === String(requestId)) { request = existingRequests[i]; break; }
    }
    if (!request) return { success: false, error: 'request_not_found' };
    if (request.status !== 'pending') return { success: false, error: 'not_pending' };

    var sSh = _ensureExactSheet('Students', H.Students);
    var existingStudents = _readFresh('Students', H.Students);
    var student = null;
    for (var si = 0; si < existingStudents.length; si++) {
      if (String(existingStudents[si].id) === String(request.studentId)) { student = existingStudents[si]; break; }
    }
    if (!student) return { success: false, error: 'student_not_found' };

    var points = Number(request.points) || 0;
    if ((Number(student.score) || 0) < points) return { success: false, error: 'insufficient_points' };

    var rSh = _ensureExactSheet('Rewards', H.Rewards);
    var existingRewards = _readFresh('Rewards', H.Rewards);
    var reward = null;
    for (var ri = 0; ri < existingRewards.length; ri++) {
      if (String(existingRewards[ri].id) === String(request.rewardId)) { reward = existingRewards[ri]; break; }
    }
    if (!reward) return { success: false, error: 'reward_not_found' };
    if ((Number(reward.quantity) || 0) <= 0) return { success: false, error: 'insufficient_reward' };

    student.score = (Number(student.score) || 0) - points;
    _upsertNoLock(sSh, 'Students', H.Students, existingStudents, [student], 'id');

    request.status = 'approved';
    request.approveDate = _nowLocale();
    _upsertNoLock(exSh, 'ExchangeHistory', H.ExchangeHistory, existingRequests, [request], 'id');

    reward.quantity = Math.max(0, (Number(reward.quantity) || 0) - 1);
    _upsertNoLock(rSh, 'Rewards', H.Rewards, existingRewards, [reward], 'id');

    var hist = {
      id: _makeHistoryId(),
      type: 'exchange',
      targetId: student.id,
      targetName: student.name,
      groupName: student.groupName || '',
      scoreChange: -points,
      reason: '兌換獎品：' + (request.rewardName || ''),
      date: _nowLocale(),
      affectedStudents: [{ id: student.id, name: student.name, scoreChange: -points }]
    };
    var hSh = _ensureExactSheet('ScoreHistory', H.ScoreHistory);
    _appendDirectNoLock(hSh, 'ScoreHistory', H.ScoreHistory, [hist], JSON_FIELDS.ScoreHistory);

    return { success: true, newScore: student.score };
  } catch(e) {
    _logError('atomicApproveExchange', e.toString());
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

// Atomically reject an exchange request (no score changes involved).
function atomicRejectExchange(requestId) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(e) {
    _logError('atomicRejectExchange-lock', e);
    return { success: false, error: 'lock_timeout' };
  }
  try {
    var exSh = _ensureExactSheet('ExchangeHistory', H.ExchangeHistory);
    var existingRequests = _readFresh('ExchangeHistory', H.ExchangeHistory);
    var request = null;
    for (var i = 0; i < existingRequests.length; i++) {
      if (String(existingRequests[i].id) === String(requestId)) { request = existingRequests[i]; break; }
    }
    if (!request) return { success: false, error: 'request_not_found' };
    if (request.status !== 'pending') return { success: false, error: 'not_pending' };

    request.status = 'rejected';
    request.rejectDate = _nowLocale();
    _upsertNoLock(exSh, 'ExchangeHistory', H.ExchangeHistory, existingRequests, [request], 'id');

    return { success: true };
  } catch(e) {
    _logError('atomicRejectExchange', e.toString());
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

// Batch read for frontend refresh after atomic operations (reduces round-trips).
function getScoreRelatedData() {
  try {
    return {
      students: _getAllObjects('Students', H.Students),
      groups: _getAllObjects('Groups', H.Groups),
      scoreHistory: _getAllObjects('ScoreHistory', H.ScoreHistory, JSON_FIELDS.ScoreHistory),
      quizAnswers: _getAllObjects('QuizAnswers', H.QuizAnswers),
      exchangeRequests: _getAllObjects('ExchangeHistory', H.ExchangeHistory),
      rewards: _getAllObjects('Rewards', H.Rewards)
    };
  } catch(e) {
    _logError('getScoreRelatedData', e.toString());
    throw e;
  }
}
