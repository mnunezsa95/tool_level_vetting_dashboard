function vetLevels() {
  const ui = SpreadsheetApp.getUi();
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet();
  const urlPattern = /^(https?:\/\/[^\s]+)/;
  let timetableURL;
  let calendarURL;
  let timetableSpreadsheet;
  let calendarSpreadsheet;
  const frequencyDays = [];
  const frequencyMapping = [];
  let possibleLessonSlotsByTerm = [];
  let totalLessons = 0;

  const programName = activeSheet.getRange("B9").getValue().trim();
  const courseName = activeSheet.getRange("C9").getValue().trim();
  const gradeLevel = String(activeSheet.getRange("D9").getValue()).trim();
  const levels = activeSheet.getRange("E9").getValue().trim();
  const isOnRamp = activeSheet.getRange("F9").getValue() === "Yes";
  const accelerationType = activeSheet.getRange("G9").getValue().trim();
  let frequency = activeSheet.getRange("H9").getValue().trim();

  let possibleLessonSlots = activeSheet.getRange("I9").getValue().trim();

  const gradePrefixMap = {
    BayelsaPRIME: "P",
    "Bridge Andhra Pradesh": "S",
    "Bridge Kenya": "G",
    "Bridge Liberia": "P",
    "Bridge Nigeria": "P",
    "Bridge Uganda": "G",
    EdoBEST: "P",
    "EdoBEST JSS": "JSS",
    EKOEXCEL: "P",
    "ESPOIR République Centrafricaine": "",
    KwaraLEARN: "P",
    RwandaEQUIP: "P",
    "STAR Education": "C",
  };

  const gradePrefix = gradePrefixMap[programName] || "";

  const dayOrder = ["M", "T", "W", "Th", "F"];
  const dayAbbreviations = {
    Monday: "M",
    Tuesday: "T",
    Wednesday: "W",
    Thursday: "Th",
    Friday: "F",
  };

  const levelData = {
    "English Studies - Reading": {
      A: { numberOfLessons: 135 },
      B: { numberOfLessons: 135 },
      C: { numberOfLessons: 135 },
      D: { numberOfLessons: 135 },
    },
    Mathematics: {
      A: { numberOfLessons: 135 },
      B: { numberOfLessons: 135 },
      C: { numberOfLessons: 150 },
      D: { numberOfLessons: 135 },
      E: { numberOfLessons: 135 },
    },
    "English Studies - Language": {
      A: { numberOfLessons: 130 },
      B: { numberOfLessons: 130 },
      C: { numberOfLessons: 130 },
      D: { numberOfLessons: 130 },
      E: { numberOfLessons: 130 },
    },
    Science: {
      A: { numberOfLessons: 100 },
      B: { numberOfLessons: 100 },
      C: { numberOfLessons: 100 },
      D: { numberOfLessons: 100 },
      E: { numberOfLessons: 100 },
    },
  };

  const courseLevels = levelData[courseName];

  if (levels.includes("-")) {
    const [startLevel, endLevel] = levels.split("-");
    const startChar = startLevel.charAt(0);
    const endChar = endLevel.charAt(0);

    for (let level of Object.keys(courseLevels)) {
      if (level >= startChar && level <= endChar) {
        totalLessons += courseLevels[level].numberOfLessons;
      }
    }
  } else if (levels) {
    const levelChar = levels.charAt(0);
    if (courseLevels[levelChar]) {
      totalLessons += courseLevels[levelChar].numberOfLessons;
    }
  }

  if (isOnRamp) {
    totalLessons = totalLessons + 20;
  }

  if (urlPattern.test(frequency)) {
    timetableURL = frequency;
    frequency = 0;
    timetableSpreadsheet = SpreadsheetApp.openByUrl(timetableURL);
    const timetableSheets = timetableSpreadsheet.getSheets();
    let timetableSheet = timetableSheets.find((sheet) => sheet.getName().includes(gradeLevel)).getName();

    const columnRowData = ui
      .prompt(
        `Enter the last column letter & last numerical row for Grade ${gradeLevel}\n Add the column then row separated by a space\n (i.e. H 18): `,
        ui.ButtonSet.OK_CANCEL
      )
      .getResponseText()
      .trim();
    let [lastColumn, lastRow] = columnRowData.split(" ");

    const currentRange = timetableSpreadsheet.getSheetByName(timetableSheet).getRange(`A1:${lastColumn}${lastRow}`);
    const values = currentRange.getValues();
    const days = values[0];

    for (let row = 1; row < values.length; row++) {
      for (let col = 0; col < values[row].length; col++) {
        const cellValue = values[row][col].toString().trim();

        if (cellValue.includes(courseName)) {
          frequency += 1;
          const day = days[col].trim();
          const abbreviatedDay = dayAbbreviations[day];
          if (abbreviatedDay && !frequencyDays.includes(abbreviatedDay)) {
            frequencyDays.push(abbreviatedDay);
          }
        }
      }
    }
  }

  // Sort frequencyDays according to dayOrder
  frequencyDays.sort((a, b) => dayOrder.indexOf(a) - dayOrder.indexOf(b));

  frequencyDays.forEach((freqValue) => {
    switch (freqValue) {
      case "M":
        frequencyMapping.push("Mon");
        break;
      case "T":
        frequencyMapping.push("Tue");
        break;
      case "W":
        frequencyMapping.push("Wed");
        break;
      case "Th":
        frequencyMapping.push("Thu");
        break;
      case "F":
        frequencyMapping.push("Fri");
        break;
      default:
        frequencyMapping.push("Unknown");
    }
  });

  if (urlPattern.test(possibleLessonSlots)) {
    calendarURL = possibleLessonSlots;
    possibleLessonSlots = 0;
    calendarSpreadsheet = SpreadsheetApp.openByUrl(calendarURL);
    const allCalendarSheets = calendarSpreadsheet.getSheets();
    const calendarSheets = [];

    for (const sheet of allCalendarSheets) {
      const sheetName = sheet.getName();
      const tempSheetName = sheetName.trim().toLowerCase();
      if (/^(t|term|s|semester)\s\d/.test(tempSheetName)) {
        calendarSheets.push(sheetName);
      }
    }

    const columnHeaderLetters = ui
      .prompt(
        `Enter the column letters for Grade ${gradeLevel}\n Add the column letter separated by a space\n (i.e. D E F): `,
        ui.ButtonSet.OK_CANCEL
      )
      .getResponseText()
      .split(" ")
      .map((char) => char.toUpperCase());

    if (columnHeaderLetters.length === calendarSheets.length) {
      calendarSheets.forEach((sheet, i) => {
        const currentSheet = calendarSpreadsheet.getSheetByName(sheet);
        const range = `${columnHeaderLetters[i]}2:${columnHeaderLetters[i]}300`;
        const targetRangeValues = currentSheet.getRange(range).getValues();

        const datesArray = currentSheet
          .getRange(`B2:B300`)
          .getValues()
          .map((date) => String(date).split(" ")[0]);

        for (let i = 0; i < targetRangeValues.length; i++) {
          let targetValue = targetRangeValues[i][0];
          let dayOfWeek = datesArray[i];

          if (frequencyMapping.includes(dayOfWeek)) {
            if (/(?:\w+\s*\/\s*\d+|\b\d+\b)/.test(targetValue)) {
              possibleLessonSlots += 1;
            }
          }
        }
        possibleLessonSlotsByTerm.push(possibleLessonSlots);
      });
    }

    possibleLessonSlotsByTerm[1] = possibleLessonSlotsByTerm[1] - possibleLessonSlotsByTerm[0];
    possibleLessonSlotsByTerm[2] =
      possibleLessonSlotsByTerm[2] - possibleLessonSlotsByTerm[1] - possibleLessonSlotsByTerm[0];

    if (courseName === "English Studies - Reading" || courseName === "Mathematics") {
      possibleLessonSlots *= 2;
      possibleLessonSlotsByTerm = possibleLessonSlotsByTerm.map((val) => val * 2);
    }
  }

  if (totalLessons <= possibleLessonSlots) {
    ui.alert(
      "Success!",
      `
      There are sufficient lesson slots to accommodate the required leveling based on the provided information. 

      Program: ${programName}
      Course: ${courseName}
      Grade: ${gradePrefix + gradeLevel}
      Level(s): ${levels}
      On-Ramp Status: ${isOnRamp === "Yes" ? true : false}
      Course Frequency: ${frequency}
      Frequency Days: ${frequencyDays.join(", ")}
      
      Total Lessons Needed: ${totalLessons}
      Possible Lesson Slots: ${possibleLessonSlots}
      Number of Possible Slots by Term/Semester: ${possibleLessonSlotsByTerm.join(", ")}
      Number of Remaining Slots: ${possibleLessonSlots - totalLessons}

      Timetable Used: ${timetableSpreadsheet.getName()}
      Calendar Used: ${calendarSpreadsheet.getName()}
      `,
      ui.ButtonSet.OK
    );
  }
}
