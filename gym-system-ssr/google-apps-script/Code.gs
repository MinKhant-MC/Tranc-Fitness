var CONFIG = {
  SPREADSHEET_ID: 'PASTE_GYM_SHEET_ID_HERE',
  SESSION_PREFIX: 'GYM_SESSION_',
  SESSION_TTL_SECONDS: 12 * 60 * 60,
  MEMBER_SESSION_TTL_SECONDS: 0,
  EXPIRY_WARNING_DAYS: 7,
  SHEETS: {
    ADMINS: 'Admins',
    REGISTRATIONS: 'Registrations',
    MEMBERS: 'Members',
    MEMBERSHIP_PAYMENTS: 'Membership Payments',
    ACTIVITIES: 'Activities',
    ATTENDANCE: 'Attendance',
    WORKOUT_LOGS: 'Workout Logs'
  },
  HEADERS: {
    ADMINS: ['admin_id', 'username', 'password_hash', 'display_name', 'status', 'created_at'],
    REGISTRATIONS: ['registration_id', 'full_name', 'phone', 'email', 'nrc', 'gender', 'weight_kg', 'height_cm', 'age', 'start_date', 'membership_months', 'personal_trainer', 'goal_note', 'photo_data', 'status', 'submitted_at', 'reviewed_at', 'reviewed_by', 'password_hash', 'membership_fee'],
    MEMBERS: ['member_id', 'full_name', 'phone', 'email', 'nrc', 'password_hash', 'gender', 'weight_kg', 'height_cm', 'age', 'start_date', 'membership_months', 'end_date', 'personal_trainer', 'goal_note', 'photo_data', 'status', 'qr_token', 'qr_code', 'approved_at', 'approved_by', 'created_at', 'membership_fee'],
    MEMBERSHIP_PAYMENTS: ['payment_id', 'member_id', 'full_name', 'phone', 'payment_type', 'months', 'amount', 'start_date', 'end_date', 'paid_at', 'created_by'],
    ACTIVITIES: ['activity_id', 'activity_name', 'category', 'calories_per_hour', 'description', 'is_active'],
    ATTENDANCE: ['attendance_id', 'member_id', 'full_name', 'phone', 'qr_code', 'scan_time', 'scan_date', 'status'],
    WORKOUT_LOGS: ['log_id', 'member_id', 'activity_id', 'activity_name', 'duration_minutes', 'estimated_calories', 'note', 'log_date', 'created_at']
  }
};

function doGet() {
  return jsonResponse_({ ok: true, data: { service: 'Cloud Share Gym API' } });
}

function doPost(e) {
  try {
    var request = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var action = cleanString_(request.action);
    var payload = request.payload || {};
    var result;

    switch (action) {
      case 'registerMember':
        result = registerMember_(payload.member || payload);
        break;
      case 'login':
        result = login_(payload);
        break;
      case 'logout':
        result = logout_(payload);
        break;
      case 'getMemberDashboard':
        result = getMemberDashboard_(requireSession_(payload, 'member'));
        break;
      case 'getAdminDashboard':
        result = getAdminDashboard_(requireSession_(payload, 'admin'), payload.query);
        break;
      case 'getGymQr':
        result = getGymQr_(requireSession_(payload));
        break;
      case 'approveRegistration':
        result = approveRegistration_(requireSession_(payload, 'admin'), payload.registration_id);
        break;
      case 'rejectRegistration':
        result = rejectRegistration_(requireSession_(payload, 'admin'), payload.registration_id);
        break;
      case 'recordAttendanceByQr':
        result = recordAttendanceByQr_(requireSession_(payload, 'admin'), payload.qr_code);
        break;
      case 'recordMemberArrival':
        result = recordMemberArrival_(requireSession_(payload, 'member'), payload.scanned_code);
        break;
      case 'recordWorkout':
        result = recordWorkout_(requireSession_(payload, 'member'), payload.workout || payload);
        break;
      case 'searchMembers':
        result = { members: searchMembers_(requireSession_(payload, 'admin'), payload.query) };
        break;
      case 'updateMember':
        result = updateMember_(requireSession_(payload, 'admin'), payload.member || payload);
        break;
      case 'removeMember':
        result = removeMember_(requireSession_(payload, 'admin'), payload.member_id);
        break;
      default:
        throw new Error('Unknown action.');
    }

    return jsonResponse_({ ok: true, data: result });
  } catch (error) {
    return jsonResponse_({ ok: false, message: error.message || 'Request failed.' });
  }
}

function setupGymManagementSystem() {
  var spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var adminSheet = getOrCreateSheet_(spreadsheet, CONFIG.SHEETS.ADMINS, CONFIG.HEADERS.ADMINS);
  var registrationSheet = getOrCreateSheet_(spreadsheet, CONFIG.SHEETS.REGISTRATIONS, CONFIG.HEADERS.REGISTRATIONS);
  var membersSheet = getOrCreateSheet_(spreadsheet, CONFIG.SHEETS.MEMBERS, CONFIG.HEADERS.MEMBERS);
  var paymentsSheet = getOrCreateSheet_(spreadsheet, CONFIG.SHEETS.MEMBERSHIP_PAYMENTS, CONFIG.HEADERS.MEMBERSHIP_PAYMENTS);
  var activitiesSheet = getOrCreateSheet_(spreadsheet, CONFIG.SHEETS.ACTIVITIES, CONFIG.HEADERS.ACTIVITIES);
  var attendanceSheet = getOrCreateSheet_(spreadsheet, CONFIG.SHEETS.ATTENDANCE, CONFIG.HEADERS.ATTENDANCE);
  var workoutSheet = getOrCreateSheet_(spreadsheet, CONFIG.SHEETS.WORKOUT_LOGS, CONFIG.HEADERS.WORKOUT_LOGS);
  var managedSheetNames = [
    CONFIG.SHEETS.ADMINS,
    CONFIG.SHEETS.REGISTRATIONS,
    CONFIG.SHEETS.MEMBERS,
    CONFIG.SHEETS.MEMBERSHIP_PAYMENTS,
    CONFIG.SHEETS.ACTIVITIES,
    CONFIG.SHEETS.ATTENDANCE,
    CONFIG.SHEETS.WORKOUT_LOGS
  ];

  if (adminSheet.getLastRow() < 2) {
    appendRecord_(adminSheet, CONFIG.HEADERS.ADMINS, {
      admin_id: 'A0001',
      username: 'admin',
      password_hash: hashText_('admin123'),
      display_name: 'Gym Admin',
      status: 'active',
      created_at: nowString_()
    });
  }

  syncDefaultActivities_(activitiesSheet);
  backfillMembershipPayments_(membersSheet, paymentsSheet);

  removeUnusedDefaultSheets_(spreadsheet, managedSheetNames);
  SpreadsheetApp.flush();
  Logger.log('Gym setup complete for spreadsheet: ' + spreadsheet.getName() + ' (' + spreadsheet.getId() + ')');
  Logger.log('Created tabs: ' + managedSheetNames.join(', '));

  return {
    message: 'Gym management system is ready.',
    spreadsheet_id: spreadsheet.getId(),
    spreadsheet_name: spreadsheet.getName(),
    sheet_names: spreadsheet.getSheets().map(function (sheet) {
      return sheet.getName();
    })
  };
}

function syncDefaultActivities_(activitiesSheet) {
  var existing = readTable_(activitiesSheet, CONFIG.HEADERS.ACTIVITIES);
  var known = {};
  var defaults = [
    ['M001', 'Treadmill Machine', '1 Cardio Machines', 600, 'Speed, incline and time', 'TRUE'],
    ['M002', 'Spinning Bike', '1 Cardio Machines', 520, 'Bike speed/resistance and time', 'TRUE'],
    ['M003', 'Power Rack', '2 Strength Machines', 430, 'Rack exercise: kg, reps and sets', 'TRUE'],
    ['M004', 'Chest Press Machine', '2 Strength Machines', 420, 'Chest press: kg, reps and sets', 'TRUE'],
    ['M005', 'Shoulder Press Machine', '2 Strength Machines', 380, 'Shoulder press: kg, reps and sets', 'TRUE'],
    ['M006', 'Leg Press Machine', '2 Strength Machines', 430, 'Leg press: kg, reps and sets', 'TRUE'],
    ['M007', 'Leg Extension Machine', '2 Strength Machines', 360, 'Leg extension: kg, reps and sets', 'TRUE'],
    ['M008', 'Leg Curls Machine', '2 Strength Machines', 360, 'Leg curls: kg, reps and sets', 'TRUE'],
    ['M009', 'Lat Pulldown Machine', '2 Strength Machines', 400, 'Lat pulldown: kg, reps and sets', 'TRUE'],
    ['M010', 'Smith Machine', '2 Strength Machines', 450, 'Smith machine: kg, reps and sets', 'TRUE'],
    ['M011', 'Hip Abduction Machine', '2 Strength Machines', 330, 'Hip abduction: kg, reps and sets', 'TRUE'],
    ['M012', 'Hip Thrust Machine', '2 Strength Machines', 390, 'Hip thrust: kg, reps and sets', 'TRUE'],
    ['M013', 'T-Bar Row', '2 Strength Machines', 420, 'T-bar row: kg, reps and sets', 'TRUE'],
    ['M014', 'Cable Machine', '2 Strength Machines', 370, 'Cable exercise: kg, reps and sets', 'TRUE'],
    ['M015', 'Seated Row Machine', '2 Strength Machines', 420, 'Seated row: kg, reps and sets', 'TRUE'],
    ['M016', 'Low Row Machine', '2 Strength Machines', 410, 'Low row: kg, reps and sets', 'TRUE'],
    ['M017', 'Hack Squat Machine', '2 Strength Machines', 440, 'Hack squat: kg, reps and sets', 'TRUE'],
    ['M018', 'Bicep Curl Machine', '2 Strength Machines', 320, 'Bicep curl: kg, reps and sets', 'TRUE'],
    ['M019', 'Push Up', '3 Movement Exercises', 300, 'Bodyweight movement: reps and sets', 'TRUE'],
    ['M020', 'Squat', '3 Movement Exercises', 360, 'Bodyweight or weighted movement: reps and sets', 'TRUE'],
    ['M021', 'Lunges', '3 Movement Exercises', 340, 'Bodyweight or weighted movement: reps and sets', 'TRUE'],
    ['M022', 'Plank', '3 Movement Exercises', 240, 'Core hold movement', 'TRUE'],
    ['M023', 'Burpees', '3 Movement Exercises', 500, 'Full body movement: reps and sets', 'TRUE'],
    ['M024', 'Mountain Climbers', '3 Movement Exercises', 420, 'Core cardio movement: reps and sets', 'TRUE'],
    ['M025', 'Jumping Jacks', '3 Movement Exercises', 380, 'Cardio movement: reps and sets', 'TRUE'],
    ['M026', 'Sit Up', '3 Movement Exercises', 280, 'Core movement: reps and sets', 'TRUE']
  ];
  var legacyNames = {
    'cycling': true,
    'strength training': true,
    'rowing': true,
    'yoga': true,
    'hiit': true,
    'chest press': true,
    'pec fly': true,
    'shoulder press': true,
    'lat pulldown': true,
    'seated row': true,
    'leg press': true,
    'leg extension': true,
    'leg curl': true,
    'inner thigh adductor': true,
    'hack squat leg press': true,
    'pec deck / chest fly machine': true,
    'cable crossover machine': true,
    'incline chest press machine': true,
    'assisted pull-up machine': true,
    'back extension machine': true,
    'lateral raise machine': true,
    'rear delt fly machine': true,
    'biceps curl machine': true,
    'triceps extension machine': true,
    'seated leg curl machine': true,
    'lying leg curl machine': true,
    'hip abductor machine': true,
    'hip adductor machine': true,
    'calf raise machine': true,
    'ab crunch machine': true,
    'rotary torso machine': true,
    'captain\'s chair / leg raise station': true,
    'treadmill': true,
    'elliptical machine': true,
    'stationary bike': true,
    'stair climber': true,
    'rowing machine': true
  };
  var desiredByName = {};
  var caloriesColumn = CONFIG.HEADERS.ACTIVITIES.indexOf('calories_per_hour') + 1;
  var categoryColumn = CONFIG.HEADERS.ACTIVITIES.indexOf('category') + 1;
  var descriptionColumn = CONFIG.HEADERS.ACTIVITIES.indexOf('description') + 1;
  var activeColumn = CONFIG.HEADERS.ACTIVITIES.indexOf('is_active') + 1;

  defaults.forEach(function (row) {
    desiredByName[String(row[1]).toLowerCase()] = row;
  });

  existing.forEach(function (item) {
    var desired = desiredByName[String(item.activity_name || '').toLowerCase()];
    known[String(item.activity_id || '').toLowerCase()] = true;
    known[String(item.activity_name || '').toLowerCase()] = true;
    if (desired) {
      if (categoryColumn > 0) {
        activitiesSheet.getRange(item._rowNumber, categoryColumn).setValue(desired[2]);
      }
      if (caloriesColumn > 0) {
        activitiesSheet.getRange(item._rowNumber, caloriesColumn).setValue(desired[3]);
      }
      if (descriptionColumn > 0) {
        activitiesSheet.getRange(item._rowNumber, descriptionColumn).setValue(desired[4]);
      }
      if (activeColumn > 0) {
        activitiesSheet.getRange(item._rowNumber, activeColumn).setValue('TRUE');
      }
    }
    if (legacyNames[String(item.activity_name || '').toLowerCase()] && activeColumn > 0) {
      activitiesSheet.getRange(item._rowNumber, activeColumn).setValue('FALSE');
    }
    if (
      String(item.activity_id || '').toUpperCase().charAt(0) === 'M' &&
      !desiredByName[String(item.activity_name || '').toLowerCase()] &&
      activeColumn > 0
    ) {
      activitiesSheet.getRange(item._rowNumber, activeColumn).setValue('FALSE');
    }
  });

  defaults.forEach(function (row) {
    if (known[String(row[0]).toLowerCase()] || known[String(row[1]).toLowerCase()]) {
      return;
    }
    activitiesSheet.appendRow(row);
  });
}

function registerMember_(member) {
  var sheet = getSheet_(CONFIG.SHEETS.REGISTRATIONS, CONFIG.HEADERS.REGISTRATIONS);
  var existingPending = findPendingRegistrationByPhone_(member.phone);
  var existingMember = findMemberByPhone_(member.phone);
  var registrationId;
  var record;

  if (existingMember && cleanString_(existingMember.status).toLowerCase() === 'active') {
    throw new Error('This phone number is already registered as a member.');
  }

  registrationId = existingPending && existingPending.registration_id
    ? existingPending.registration_id
    : nextId_(sheet, 'registration_id', 'R');

  record = {
    registration_id: registrationId,
    full_name: requiredString_(member.full_name, 'Name is required.'),
    phone: requiredString_(member.phone, 'Phone is required.'),
    email: cleanString_(member.email),
    nrc: cleanString_(member.nrc),
    gender: requiredString_(member.gender, 'Gender is required.'),
    weight_kg: positiveNumber_(member.weight_kg, 'Weight is invalid.'),
    height_cm: positiveNumber_(member.height_cm, 'Height is invalid.'),
    age: positiveNumber_(member.age, 'Age is invalid.'),
    start_date: normalizeDate_(member.start_date),
    membership_months: positiveNumber_(member.membership_months, 'Membership months are invalid.'),
    personal_trainer: cleanString_(member.personal_trainer) || 'No',
    goal_note: cleanString_(member.goal_note),
    photo_data: requiredString_(member.photo_data, 'Member photo is required.'),
    status: 'pending',
    submitted_at: nowString_(),
    reviewed_at: '',
    reviewed_by: '',
    password_hash: hashText_(requiredString_(member.password, 'Password is required.')),
    membership_fee: positiveNumber_(member.membership_fee || member.income, 'Membership fee is invalid.')
  };

  if (existingPending && existingPending._rowNumber) {
    writeRecord_(sheet, existingPending._rowNumber, CONFIG.HEADERS.REGISTRATIONS, record);
  } else {
    appendRecord_(sheet, CONFIG.HEADERS.REGISTRATIONS, record);
  }

  return {
    message: existingPending
      ? 'Pending registration updated. Waiting for admin approval.'
      : 'Registration submitted. Waiting for admin approval.',
    registration_id: registrationId,
    status: 'pending'
  };
}

function login_(payload) {
  var role = cleanString_(payload.role).toLowerCase() || 'member';
  var identifier = requiredString_(payload.identifier, 'Identifier is required.');
  var password = requiredString_(payload.password, 'Password is required.');
  var token = Utilities.getUuid();
  var loginData;

  if (role === 'admin') {
    loginData = loginAdmin_(identifier, password);
  } else {
    loginData = loginMember_(identifier, password);
  }

  PropertiesService.getScriptProperties().setProperty(CONFIG.SESSION_PREFIX + token, JSON.stringify({
    token: token,
    role: loginData.role,
    user_id: loginData.user_id,
    expires_at: buildSessionExpiry_(loginData.role)
  }));

  return {
    session_token: token,
    role: loginData.role,
    user_id: loginData.user_id,
    display_name: loginData.display_name
  };
}

function logout_(payload) {
  var token = cleanString_(payload.session_token);
  if (token) {
    PropertiesService.getScriptProperties().deleteProperty(CONFIG.SESSION_PREFIX + token);
  }
  return { success: true };
}

function loginAdmin_(identifier, password) {
  var admins = readTable_(getSheet_(CONFIG.SHEETS.ADMINS, CONFIG.HEADERS.ADMINS), CONFIG.HEADERS.ADMINS);
  var match = admins.filter(function (admin) {
    return cleanString_(admin.username).toLowerCase() === cleanString_(identifier).toLowerCase();
  })[0];

  if (!match || cleanString_(match.status).toLowerCase() !== 'active') {
    throw new Error('Admin account not found.');
  }
  if (!verifyPassword_(password, match.password_hash)) {
    throw new Error('Admin password is incorrect.');
  }

  return {
    role: 'admin',
    user_id: match.admin_id,
    display_name: match.display_name || match.username
  };
}

function loginMember_(identifier, password) {
  var members = readTable_(getSheet_(CONFIG.SHEETS.MEMBERS, CONFIG.HEADERS.MEMBERS), CONFIG.HEADERS.MEMBERS);
  var match = members.filter(function (member) {
    return normalizePhone_(member.phone) === normalizePhone_(identifier);
  })[0];

  if (!match || cleanString_(match.status).toLowerCase() !== 'active') {
    throw new Error('Member is not active or not yet approved.');
  }
  if (!verifyPassword_(password, match.password_hash)) {
    throw new Error('Member password is incorrect.');
  }

  return {
    role: 'member',
    user_id: match.member_id,
    display_name: match.full_name
  };
}

function requireSession_(payload, expectedRole) {
  var token = requiredString_(payload.session_token, 'Session token missing.');
  var raw = PropertiesService.getScriptProperties().getProperty(CONFIG.SESSION_PREFIX + token);
  var session;

  if (!raw) {
    throw new Error('Session expired. Please log in again.');
  }

  session = JSON.parse(raw);
  if (Number(session.expires_at) > 0 && session.expires_at < Date.now()) {
    PropertiesService.getScriptProperties().deleteProperty(CONFIG.SESSION_PREFIX + token);
    throw new Error('Session expired. Please log in again.');
  }
  if (expectedRole && session.role !== expectedRole) {
    throw new Error('Access denied for this account type.');
  }
  return session;
}

function buildSessionExpiry_(role) {
  if (cleanString_(role).toLowerCase() === 'member' && Number(CONFIG.MEMBER_SESSION_TTL_SECONDS) === 0) {
    return 0;
  }

  if (cleanString_(role).toLowerCase() === 'member') {
    return Date.now() + (Number(CONFIG.MEMBER_SESSION_TTL_SECONDS) * 1000);
  }

  return Date.now() + (CONFIG.SESSION_TTL_SECONDS * 1000);
}

function getMemberDashboard_(session) {
  var member = findMemberById_(session.user_id);
  var activities = getActiveActivities_();
  var sevenDayDates;
  var logsSevenDays;
  var logsToday;
  var sevenDaySummary;
  var sevenDayCalories;
  var checkedInToday = Boolean(findTodayAttendanceByMemberId_(member.member_id));
  var caloriesToday;
  var heightMeters = toNumber_(member.height_cm) / 100;
  var bmi = heightMeters > 0 ? round2_(toNumber_(member.weight_kg) / (heightMeters * heightMeters)) : 0;

  cleanupOldWorkoutLogs_();
  sevenDayDates = getLastSevenDates_();
  logsSevenDays = getWorkoutLogsForMemberRange_(member.member_id, sevenDayDates);
  logsToday = logsSevenDays.filter(function (item) {
    return cleanString_(item.log_date) === todayString_();
  });
  sevenDaySummary = buildSevenDaySummary_(logsSevenDays, sevenDayDates);
  caloriesToday = logsToday.reduce(function (sum, item) {
    return sum + toNumber_(item.estimated_calories);
  }, 0);
  sevenDayCalories = sevenDaySummary.reduce(function (sum, item) {
    return sum + toNumber_(item.calories);
  }, 0);

  return {
    member: member,
    bmi: bmi,
    calories_today: round2_(caloriesToday),
    calories_7_days: round2_(sevenDayCalories),
    checked_in_today: checkedInToday,
    gym_checkin_qr: getGymCheckinQrValue_(),
    activities: activities,
    workout_logs_today: logsToday,
    workout_logs_7_days: logsSevenDays,
    seven_day_activity_summary: sevenDaySummary,
    notifications: getMemberNotifications_(member)
  };
}

function getAdminDashboard_(session, query) {
  cleanupOldMemberships_();
  var pending = readTable_(getSheet_(CONFIG.SHEETS.REGISTRATIONS, CONFIG.HEADERS.REGISTRATIONS), CONFIG.HEADERS.REGISTRATIONS)
    .filter(function (item) { return cleanString_(item.status).toLowerCase() === 'pending'; });
  var attendance = getTodayAttendance_();
  var members = getActiveMembers_();
  var payments = getMembershipPayments_();
  var monthlyIncome = buildMonthlyFeeSummary_(payments);
  var totalIncome = monthlyIncome.reduce(function (sum, item) {
    return sum + toNumber_(item.amount);
  }, 0);
  var expiring = members.filter(function (member) {
    return daysUntil_(member.end_date) <= CONFIG.EXPIRY_WARNING_DAYS;
  });

  return {
    stats: {
      total_members: members.length,
      pending_count: pending.length,
      present_today: attendance.length,
      expiring_soon: expiring.length,
      total_income: totalIncome
    },
    monthly_income: monthlyIncome,
    payment_history: payments.slice().reverse(),
    pending_registrations: pending,
    all_members: members,
    today_attendance: attendance,
    gym_checkin_qr: getGymCheckinQrValue_(),
    notifications: expiring.map(function (member) {
      return {
        title: 'Membership expiring soon',
        message: member.full_name + ' expires in ' + daysUntil_(member.end_date) + ' day(s).'
      };
    }),
    search_results: searchMembers_(session, query)
  };
}

function approveRegistration_(session, registrationId) {
  var registrationsSheet = getSheet_(CONFIG.SHEETS.REGISTRATIONS, CONFIG.HEADERS.REGISTRATIONS);
  var membersSheet = getSheet_(CONFIG.SHEETS.MEMBERS, CONFIG.HEADERS.MEMBERS);
  var found = findRecordById_(registrationsSheet, CONFIG.HEADERS.REGISTRATIONS, 'registration_id', registrationId);
  var registration;
  var memberId;
  var qrToken;

  if (!found) {
    throw new Error('Registration not found.');
  }

  registration = found.record;
  if (cleanString_(registration.status).toLowerCase() !== 'pending') {
    throw new Error('This registration has already been reviewed.');
  }

  memberId = nextId_(membersSheet, 'member_id', 'M');
  qrToken = Utilities.getUuid().replace(/-/g, '').substring(0, 16);

  appendRecord_(membersSheet, CONFIG.HEADERS.MEMBERS, {
    member_id: memberId,
    full_name: registration.full_name,
    phone: registration.phone,
    email: registration.email,
    nrc: cleanString_(registration.nrc),
    password_hash: cleanString_(registration.password_hash) || hashText_(registration.phone),
    gender: registration.gender,
    weight_kg: registration.weight_kg,
    height_cm: registration.height_cm,
    age: registration.age,
    start_date: registration.start_date,
    membership_months: registration.membership_months,
    end_date: addMonths_(registration.start_date, registration.membership_months),
    personal_trainer: registration.personal_trainer,
    goal_note: registration.goal_note,
    photo_data: registration.photo_data,
    status: 'active',
    qr_token: qrToken,
    qr_code: 'GYM|' + memberId + '|' + qrToken,
    approved_at: nowString_(),
    approved_by: session.user_id,
    created_at: nowString_(),
    membership_fee: registration.membership_fee
  });

  recordMembershipPayment_({
    member_id: memberId,
    full_name: registration.full_name,
    phone: registration.phone,
    payment_type: 'registration',
    months: registration.membership_months,
    amount: registration.membership_fee,
    start_date: registration.start_date,
    end_date: addMonths_(registration.start_date, registration.membership_months),
    paid_at: nowString_(),
    created_by: session.user_id
  });

  registration.status = 'approved';
  registration.reviewed_at = nowString_();
  registration.reviewed_by = session.user_id;
  writeRecord_(registrationsSheet, found.rowNumber, CONFIG.HEADERS.REGISTRATIONS, registration);

  return { success: true };
}

function getGymQr_(session) {
  return {
    gym_checkin_qr: getGymCheckinQrValue_(),
    role: session.role
  };
}

function rejectRegistration_(session, registrationId) {
  var sheet = getSheet_(CONFIG.SHEETS.REGISTRATIONS, CONFIG.HEADERS.REGISTRATIONS);
  var found = findRecordById_(sheet, CONFIG.HEADERS.REGISTRATIONS, 'registration_id', registrationId);

  if (!found) {
    throw new Error('Registration not found.');
  }

  found.record.status = 'rejected';
  found.record.reviewed_at = nowString_();
  found.record.reviewed_by = session.user_id;
  writeRecord_(sheet, found.rowNumber, CONFIG.HEADERS.REGISTRATIONS, found.record);
  return { success: true };
}

function recordAttendanceByQr_(session, qrCode) {
  var membersSheet = getSheet_(CONFIG.SHEETS.MEMBERS, CONFIG.HEADERS.MEMBERS);
  var members = readTable_(membersSheet, CONFIG.HEADERS.MEMBERS);
  var member = members.filter(function (item) {
    return cleanString_(item.qr_code) === cleanString_(qrCode) && cleanString_(item.status).toLowerCase() === 'active';
  })[0];

  if (!member) {
    throw new Error('QR code not recognized.');
  }

  return recordAttendanceForMember_(member, member.qr_code);
}

function recordMemberArrival_(session, scannedCode) {
  var member = findMemberById_(session.user_id);
  var expectedGymCode = getGymCheckinQrValue_();

  if (cleanString_(scannedCode) !== cleanString_(expectedGymCode)) {
    throw new Error('This is not the gym check-in QR code.');
  }

  return recordAttendanceForMember_(member, expectedGymCode);
}

function recordWorkout_(session, workout) {
  var member = findMemberById_(session.user_id);
  var activity = findActivityById_(workout.activity_id);
  var duration = positiveNumber_(workout.duration_minutes, 'Duration is invalid.');
  var estimatedCalories;
  var sheet = getSheet_(CONFIG.SHEETS.WORKOUT_LOGS, CONFIG.HEADERS.WORKOUT_LOGS);

  if (!activity) {
    throw new Error('Activity not found.');
  }

  estimatedCalories = toNumber_(workout.estimated_calories) > 0
    ? round2_(toNumber_(workout.estimated_calories))
    : round2_(toNumber_(activity.calories_per_hour) * (duration / 60));

  appendRecord_(sheet, CONFIG.HEADERS.WORKOUT_LOGS, {
    log_id: nextId_(sheet, 'log_id', 'L'),
    member_id: member.member_id,
    activity_id: activity.activity_id,
    activity_name: activity.activity_name,
    duration_minutes: duration,
    estimated_calories: estimatedCalories,
    note: cleanString_(workout.note),
    log_date: todayString_(),
    created_at: nowString_()
  });

  cleanupOldWorkoutLogs_();

  return { success: true };
}

function updateMember_(session, memberData) {
  var sheet = getSheet_(CONFIG.SHEETS.MEMBERS, CONFIG.HEADERS.MEMBERS);
  var memberId = requiredString_(memberData.member_id, 'Member ID is required.');
  var found = findRecordById_(sheet, CONFIG.HEADERS.MEMBERS, 'member_id', memberId);
  var record;
  var currentPasswordHash;
  var extendMonths;
  var extensionFee;
  var extensionStartDate;

  if (!found) {
    throw new Error('Member not found.');
  }

  record = found.record;
  currentPasswordHash = record.password_hash;
  [
    'full_name',
    'phone',
    'email',
    'gender',
    'weight_kg',
    'height_cm',
    'age',
    'start_date',
    'membership_months',
    'personal_trainer',
    'goal_note',
    'status'
  ].forEach(function (key) {
    if (Object.prototype.hasOwnProperty.call(memberData, key)) {
      record[key] = cleanString_(memberData[key]);
    }
  });

  record.full_name = requiredString_(record.full_name, 'Name is required.');
  record.phone = requiredString_(record.phone, 'Phone is required.');
  record.gender = requiredString_(record.gender, 'Gender is required.');
  record.weight_kg = positiveNumber_(record.weight_kg, 'Weight is invalid.');
  record.height_cm = positiveNumber_(record.height_cm, 'Height is invalid.');
  record.age = positiveNumber_(record.age, 'Age is invalid.');
  record.start_date = normalizeDate_(record.start_date);
  record.membership_months = positiveNumber_(record.membership_months, 'Membership months are invalid.');
  record.membership_fee = toNumber_(record.membership_fee);
  extendMonths = Number(memberData.extend_months || 0);
  extensionFee = toNumber_(memberData.extension_fee);
  if (extensionFee > 0 && extendMonths <= 0) {
    throw new Error('Extend months is required when adding an extension fee.');
  }
  if (extendMonths > 0) {
    extensionStartDate = record.end_date || addMonths_(record.start_date, record.membership_months);
    record.membership_months = Number(record.membership_months) + extendMonths;
  }
  record.personal_trainer = cleanString_(record.personal_trainer) || 'No';
  record.status = cleanString_(record.status) || 'active';
  record.end_date = addMonths_(record.start_date, record.membership_months);

  if (extendMonths > 0 && extensionFee > 0) {
    record.membership_fee = toNumber_(record.membership_fee) + extensionFee;
    recordMembershipPayment_({
      member_id: record.member_id,
      full_name: record.full_name,
      phone: record.phone,
      payment_type: 'extension',
      months: extendMonths,
      amount: extensionFee,
      start_date: extensionStartDate,
      end_date: record.end_date,
      paid_at: nowString_(),
      created_by: session.user_id
    });
  }

  if (cleanString_(memberData.new_password)) {
    record.password_hash = hashText_(requiredString_(memberData.new_password, 'New password is required.'));
  } else if (!currentPasswordHash) {
    record.password_hash = hashText_(record.phone);
  } else {
    record.password_hash = currentPasswordHash;
  }

  writeRecord_(sheet, found.rowNumber, CONFIG.HEADERS.MEMBERS, record);
  return { success: true, member: record };
}

function removeMember_(session, memberId) {
  var sheet = getSheet_(CONFIG.SHEETS.MEMBERS, CONFIG.HEADERS.MEMBERS);
  var found = findRecordById_(sheet, CONFIG.HEADERS.MEMBERS, 'member_id', memberId);

  if (!found) {
    throw new Error('Member not found.');
  }

  found.record.status = 'removed';
  writeRecord_(sheet, found.rowNumber, CONFIG.HEADERS.MEMBERS, found.record);
  return { success: true };
}

function searchMembers_(session, query) {
  var members = getActiveMembers_();
  var target = cleanString_(query).toLowerCase();
  var normalizedTarget = normalizePhone_(query);

  if (!target) {
    return members.slice(0, 20);
  }

  return members.filter(function (member) {
    return [member.full_name, member.phone, member.member_id].join(' ').toLowerCase().indexOf(target) !== -1 ||
      (normalizedTarget && normalizePhone_(member.phone).indexOf(normalizedTarget) !== -1);
  });
}

function getTodayAttendance_() {
  return readTable_(getSheet_(CONFIG.SHEETS.ATTENDANCE, CONFIG.HEADERS.ATTENDANCE), CONFIG.HEADERS.ATTENDANCE)
    .filter(function (item) { return cleanString_(item.scan_date) === todayString_(); });
}

function findTodayAttendanceByMemberId_(memberId) {
  return readTable_(getSheet_(CONFIG.SHEETS.ATTENDANCE, CONFIG.HEADERS.ATTENDANCE), CONFIG.HEADERS.ATTENDANCE)
    .filter(function (item) {
      return cleanString_(item.member_id) === cleanString_(memberId) && cleanString_(item.scan_date) === todayString_();
    })[0] || null;
}

function recordAttendanceForMember_(member, qrCode) {
  var attendanceSheet = getSheet_(CONFIG.SHEETS.ATTENDANCE, CONFIG.HEADERS.ATTENDANCE);
  var existing = findTodayAttendanceByMemberId_(member.member_id);

  if (existing) {
    throw new Error('This member has already checked in today.');
  }

  appendRecord_(attendanceSheet, CONFIG.HEADERS.ATTENDANCE, {
    attendance_id: nextId_(attendanceSheet, 'attendance_id', 'AT'),
    member_id: member.member_id,
    full_name: member.full_name,
    phone: member.phone,
    qr_code: qrCode,
    scan_time: nowString_(),
    scan_date: todayString_(),
    status: 'present'
  });

  return { success: true, checked_in_today: true };
}

function getGymCheckinQrValue_() {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(CONFIG.SPREADSHEET_ID || 'TRANC_GYM'));
  var token = bytes.map(function (value) {
    var normalized = value < 0 ? value + 256 : value;
    return ('0' + normalized.toString(16)).slice(-2);
  }).join('').substring(0, 24).toUpperCase();

  return 'TRANC_GYM|ENTRY|' + token;
}

function getWorkoutLogsForMember_(memberId, date) {
  return readTable_(getSheet_(CONFIG.SHEETS.WORKOUT_LOGS, CONFIG.HEADERS.WORKOUT_LOGS), CONFIG.HEADERS.WORKOUT_LOGS)
    .filter(function (item) {
      return cleanString_(item.member_id) === cleanString_(memberId) && cleanString_(item.log_date) === cleanString_(date);
    });
}

function getWorkoutLogsForMemberRange_(memberId, dates) {
  var allowed = {};

  dates.forEach(function (date) {
    allowed[date] = true;
  });

  return readTable_(getSheet_(CONFIG.SHEETS.WORKOUT_LOGS, CONFIG.HEADERS.WORKOUT_LOGS), CONFIG.HEADERS.WORKOUT_LOGS)
    .filter(function (item) {
      return cleanString_(item.member_id) === cleanString_(memberId) && Boolean(allowed[cleanString_(item.log_date)]);
    })
    .sort(function (a, b) {
      return cleanString_(b.log_date).localeCompare(cleanString_(a.log_date)) ||
        cleanString_(b.created_at).localeCompare(cleanString_(a.created_at));
    });
}

function buildSevenDaySummary_(logs, dates) {
  var mapped = {};

  dates.forEach(function (date) {
    mapped[date] = {
      date: date,
      calories: 0,
      duration_minutes: 0,
      activity_count: 0
    };
  });

  logs.forEach(function (log) {
    var date = cleanString_(log.log_date);
    if (!mapped[date]) {
      return;
    }
    mapped[date].calories += toNumber_(log.estimated_calories);
    mapped[date].duration_minutes += toNumber_(log.duration_minutes);
    mapped[date].activity_count += 1;
  });

  return dates.map(function (date) {
    mapped[date].calories = round2_(mapped[date].calories);
    mapped[date].duration_minutes = round2_(mapped[date].duration_minutes);
    return mapped[date];
  });
}

function getLastSevenDates_() {
  var dates = [];
  var now = new Date();
  var date;
  var index;

  for (index = 6; index >= 0; index -= 1) {
    date = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    date.setDate(date.getDate() - index);
    dates.push(Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  }

  return dates;
}

function cleanupOldWorkoutLogs_() {
  var sheet = getSheet_(CONFIG.SHEETS.WORKOUT_LOGS, CONFIG.HEADERS.WORKOUT_LOGS);
  var rows = readTable_(sheet, CONFIG.HEADERS.WORKOUT_LOGS);
  var cutoff = getLastSevenDates_()[0];
  var index;
  var date;

  for (index = rows.length - 1; index >= 0; index -= 1) {
    date = cleanString_(rows[index].log_date);
    if (date && date < cutoff) {
      sheet.deleteRow(rows[index]._rowNumber);
    }
  }
}

function cleanupOldMemberships_() {
  var sheet = getSheet_(CONFIG.SHEETS.MEMBERS, CONFIG.HEADERS.MEMBERS);
  var rows = readTable_(sheet, CONFIG.HEADERS.MEMBERS);
  var now = new Date();
  var cutoff = new Date(now.getFullYear() - 1, now.getMonth(), now.getDate());
  var index;
  var endDate;

  for (index = rows.length - 1; index >= 0; index -= 1) {
    endDate = parseSheetDate_(rows[index].end_date);
    if (endDate && endDate < cutoff) {
      sheet.deleteRow(rows[index]._rowNumber);
    }
  }
}

function getMemberNotifications_(member) {
  var days = daysUntil_(member.end_date);
  var list = [];
  if (days <= CONFIG.EXPIRY_WARNING_DAYS) {
    list.push({
      title: 'Membership expiring soon',
      message: 'Your membership will expire in ' + days + ' day(s). Please renew soon.'
    });
  }
  if (cleanString_(member.status).toLowerCase() !== 'active') {
    list.push({
      title: 'Membership not active',
      message: 'Your account is not active yet. Contact admin.'
    });
  }
  return list;
}

function getActiveMembers_() {
  return readTable_(getSheet_(CONFIG.SHEETS.MEMBERS, CONFIG.HEADERS.MEMBERS), CONFIG.HEADERS.MEMBERS)
    .filter(function (item) { return cleanString_(item.status).toLowerCase() === 'active'; });
}

function getMembershipPayments_() {
  return readTable_(getSheet_(CONFIG.SHEETS.MEMBERSHIP_PAYMENTS, CONFIG.HEADERS.MEMBERSHIP_PAYMENTS), CONFIG.HEADERS.MEMBERSHIP_PAYMENTS);
}

function recordMembershipPayment_(payment) {
  var sheet = getSheet_(CONFIG.SHEETS.MEMBERSHIP_PAYMENTS, CONFIG.HEADERS.MEMBERSHIP_PAYMENTS);
  var amount = positiveNumber_(payment.amount, 'Membership fee is invalid.');
  var months = positiveNumber_(payment.months, 'Membership months are invalid.');

  appendRecord_(sheet, CONFIG.HEADERS.MEMBERSHIP_PAYMENTS, {
    payment_id: nextId_(sheet, 'payment_id', 'F'),
    member_id: requiredString_(payment.member_id, 'Member ID is required.'),
    full_name: requiredString_(payment.full_name, 'Name is required.'),
    phone: requiredString_(payment.phone, 'Phone is required.'),
    payment_type: cleanString_(payment.payment_type) || 'registration',
    months: months,
    amount: amount,
    start_date: normalizeDate_(payment.start_date),
    end_date: normalizeDate_(payment.end_date),
    paid_at: cleanString_(payment.paid_at) || nowString_(),
    created_by: cleanString_(payment.created_by)
  });
}

function backfillMembershipPayments_(membersSheet, paymentsSheet) {
  var payments = readTable_(paymentsSheet, CONFIG.HEADERS.MEMBERSHIP_PAYMENTS);
  var members;

  if (payments.length) {
    return;
  }

  members = readTable_(membersSheet, CONFIG.HEADERS.MEMBERS);
  members.forEach(function (member) {
    var fee = toNumber_(member.membership_fee);
    if (fee <= 0 || !cleanString_(member.member_id)) {
      return;
    }
    appendRecord_(paymentsSheet, CONFIG.HEADERS.MEMBERSHIP_PAYMENTS, {
      payment_id: nextId_(paymentsSheet, 'payment_id', 'F'),
      member_id: member.member_id,
      full_name: member.full_name,
      phone: member.phone,
      payment_type: 'registration',
      months: member.membership_months,
      amount: fee,
      start_date: member.start_date,
      end_date: member.end_date,
      paid_at: cleanString_(member.approved_at) || cleanString_(member.created_at) || cleanString_(member.start_date),
      created_by: member.approved_by
    });
  });
}

function getActiveActivities_() {
  return readTable_(getSheet_(CONFIG.SHEETS.ACTIVITIES, CONFIG.HEADERS.ACTIVITIES), CONFIG.HEADERS.ACTIVITIES)
    .filter(function (item) { return cleanString_(item.is_active).toLowerCase() !== 'false'; });
}

function buildMonthlyFeeSummary_(payments) {
  var monthLabels = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var year = new Date().getFullYear();
  var months = monthLabels.map(function (label, index) {
    return {
      month: year + '-' + String(index + 1).padStart(2, '0'),
      label: label,
      amount: 0
    };
  });

  payments.forEach(function (payment) {
    var sourceDate = cleanString_(payment.paid_at) || cleanString_(payment.start_date);
    var date = parseSheetDate_(sourceDate);
    var fee = toNumber_(payment.amount);

    if (!date || date.getFullYear() !== year || fee <= 0) {
      return;
    }

    months[date.getMonth()].amount += fee;
  });

  return months;
}

function parseSheetDate_(value) {
  var date;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value;
  }
  if (!cleanString_(value)) {
    return null;
  }
  date = new Date(cleanString_(value).replace(' ', 'T'));
  return isNaN(date.getTime()) ? null : date;
}

function findActivityById_(activityId) {
  var items = getActiveActivities_();
  return items.filter(function (item) { return cleanString_(item.activity_id) === cleanString_(activityId); })[0] || null;
}

function findMemberById_(memberId) {
  var items = readTable_(getSheet_(CONFIG.SHEETS.MEMBERS, CONFIG.HEADERS.MEMBERS), CONFIG.HEADERS.MEMBERS);
  var member = items.filter(function (item) { return cleanString_(item.member_id) === cleanString_(memberId); })[0];
  if (!member) {
    throw new Error('Member not found.');
  }
  return member;
}

function findMemberByPhone_(phone) {
  var items = readTable_(getSheet_(CONFIG.SHEETS.MEMBERS, CONFIG.HEADERS.MEMBERS), CONFIG.HEADERS.MEMBERS);
  return items.filter(function (item) {
    return normalizePhone_(item.phone) === normalizePhone_(phone);
  })[0] || null;
}

function findPendingRegistrationByPhone_(phone) {
  var items = readTable_(getSheet_(CONFIG.SHEETS.REGISTRATIONS, CONFIG.HEADERS.REGISTRATIONS), CONFIG.HEADERS.REGISTRATIONS);
  var targetPhone = normalizePhone_(phone);

  if (!targetPhone) {
    return null;
  }

  return items.filter(function (item) {
    return cleanString_(item.status).toLowerCase() === 'pending' &&
      normalizePhone_(item.phone) === targetPhone;
  })[0] || null;
}

function getSheet_(name, headers) {
  var spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    throw new Error(name + ' sheet not found. Run setupGymManagementSystem first.');
  }
  ensureHeaders_(sheet, headers);
  return sheet;
}

function getOrCreateSheet_(spreadsheet, name, headers) {
  var sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
  ensureHeaders_(sheet, headers);
  return sheet;
}

function removeUnusedDefaultSheets_(spreadsheet, keepNames) {
  var sheets = spreadsheet.getSheets();

  sheets.forEach(function (sheet) {
    var name = cleanString_(sheet.getName());
    var isManaged = keepNames.indexOf(name) !== -1;
    var isDefaultBlankName = /^Sheet\d*$/i.test(name);
    var isBlank = sheet.getLastRow() === 0 && sheet.getLastColumn() === 0;

    if (!isManaged && isDefaultBlankName && isBlank && spreadsheet.getSheets().length > keepNames.length) {
      spreadsheet.deleteSheet(sheet);
    }
  });
}

function ensureHeaders_(sheet, headers) {
  var current;
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }
  current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  headers.forEach(function (header, index) {
    if (!current[index]) {
      sheet.getRange(1, index + 1).setValue(header);
    }
  });
}

function readTable_(sheet, headers) {
  var lastRow = sheet.getLastRow();
  var values;
  if (lastRow < 2) {
    return [];
  }
  values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values.map(function (row, index) {
    var item = { _rowNumber: index + 2 };
    headers.forEach(function (header, colIndex) {
      item[header] = row[colIndex];
    });
    return item;
  });
}

function appendRecord_(sheet, headers, record) {
  var row = headers.map(function (header) { return record[header] !== undefined ? record[header] : ''; });
  var nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, 1, headers.length).setNumberFormat('@');
  sheet.getRange(nextRow, 1, 1, headers.length).setValues([row]);
}

function writeRecord_(sheet, rowNumber, headers, record) {
  var row = headers.map(function (header) { return record[header] !== undefined ? record[header] : ''; });
  sheet.getRange(rowNumber, 1, 1, headers.length).setValues([row]);
}

function findRecordById_(sheet, headers, key, value) {
  var rows = readTable_(sheet, headers);
  var target = cleanString_(value);
  var found = rows.filter(function (item) {
    return cleanString_(item[key]) === target;
  })[0];
  if (!found) {
    return null;
  }
  return {
    rowNumber: found._rowNumber,
    record: found
  };
}

function nextId_(sheet, key, prefix) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var rows = readTable_(sheet, headers);
  var max = 0;
  rows.forEach(function (item) {
    var match = cleanString_(item[key]).match(/(\d+)$/);
    if (match) {
      max = Math.max(max, Number(match[1]));
    }
  });
  return prefix + String(max + 1).padStart(4, '0');
}

function hashText_(value) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, cleanString_(value), Utilities.Charset.UTF_8);
  return bytes.map(function (byte) {
    var unsigned = byte < 0 ? byte + 256 : byte;
    return ('0' + unsigned.toString(16)).slice(-2);
  }).join('');
}

function verifyPassword_(plain, hash) {
  var cleanHash = cleanString_(hash);
  return cleanHash === hashText_(plain) || cleanHash === cleanString_(plain);
}

function normalizeDate_(value) {
  var text = requiredString_(value, 'Date is required.');
  return text.substring(0, 10);
}

function addMonths_(startDate, months) {
  var date = new Date(normalizeDate_(startDate) + 'T00:00:00');
  date.setMonth(date.getMonth() + Number(months));
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function daysUntil_(dateString) {
  var target = new Date(String(dateString).substring(0, 10) + 'T00:00:00');
  var now = new Date();
  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  return Math.ceil((target.getTime() - today.getTime()) / (24 * 60 * 60 * 1000));
}

function todayString_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function nowString_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function requiredString_(value, message) {
  var text = cleanString_(value);
  if (!text) {
    throw new Error(message);
  }
  return text;
}

function positiveNumber_(value, message) {
  var number = Number(value);
  if (!Number.isFinite(number) || number <= 0) {
    throw new Error(message);
  }
  return number;
}

function toNumber_(value) {
  var number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function cleanString_(value) {
  return String(value === undefined || value === null ? '' : value).trim();
}

function normalizePhone_(value) {
  return cleanString_(value).replace(/\D/g, '').replace(/^0+/, '');
}

function round2_(value) {
  return Math.round(Number(value) * 100) / 100;
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

