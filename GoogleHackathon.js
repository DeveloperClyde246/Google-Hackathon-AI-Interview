const HTML_STYLES = `
<style>
  body {
    font-family: Arial, sans-serif;
    margin: 0;
    padding: 20px;
    background-color: #f5f5f5;
    color: #333;
  }
  h2 {
    color: #2c3e50;
    text-align: center;
    margin-bottom: 20px;
  }
  table {
    width: 100%;
    border-collapse: collapse;
    background-color: white;
    box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    margin-bottom: 20px;
  }
  th, td {
    padding: 12px;
    text-align: left;
    border-bottom: 1px solid #ddd;
  }
  th {
    background-color: #3498db;
    color: white;
  }
  tr:hover {
    background-color: #f5f5f5;
  }
  button, input[type="submit"] {
    background-color: #2ecc71;
    border: none;
    color: white;
    padding: 10px 20px;
    text-align: center;
    text-decoration: none;
    display: inline-block;
    font-size: 16px;
    margin: 4px 2px;
    cursor: pointer;
    border-radius: 4px;
    transition: background-color 0.3s;
  }
  button:hover, input[type="submit"]:hover {
    background-color: #27ae60;
  }
  input[type="text"], input[type="number"], input[type="date"], select {
    width: 100%;
    padding: 8px;
    margin: 8px 0;
    display: inline-block;
    border: 1px solid #ccc;
    border-radius: 4px;
    box-sizing: border-box;
  }
  label {
    font-weight: bold;
  }
  .form-group {
    margin-bottom: 15px;
  }
</style>
`;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('HR Interview')
    .addItem('List Candidates and Send Invites', 'listCandidatesAndSendInvites')
    .addItem('Enter Interview Scores', 'showScoreEntry')
    .addItem('Generate Feedback', 'generateFeedback')
    .addItem('Decision Support', 'showDecisionSupport')
    .addItem('Generate Report', 'generateReport')
    .addToUi();
}

function listCandidatesAndSendInvites() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var ui = SpreadsheetApp.getUi();
  
  var htmlOutput = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <base target="_top">
        ${HTML_STYLES}
      </head>
      <body>
        <h2>Candidates</h2>
        <table>
          <tr>
            <th>ID</th>
            <th>Name</th>
            <th>Position</th>
            <th>Action</th>
          </tr>
  `).setWidth(500).setHeight(400);

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    htmlOutput.append(`
          <tr>
            <td>${row[0]}</td>
            <td>${row[1]}</td>
            <td>${row[3]}</td>
            <td><button onclick="sendInvite(${i})">Send Invite</button></td>
          </tr>
    `);
  }

  htmlOutput.append(`
        </table>
        <script>
          function sendInvite(index) {
            var date = prompt("Enter interview date (YYYY-MM-DD):");
            if (date) {
              google.script.run
                .withSuccessHandler(function() {
                  alert("Invite sent successfully!");
                })
                .withFailureHandler(function(error) {
                  alert("Error: " + error.message);
                })
                .sendInterviewInvite(index, date);
            }
          }
        </script>
      </body>
    </html>
  `);

  ui.showModalDialog(htmlOutput, 'Send Interview Invites');
}

function sendInterviewInvite(index, interviewDate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  var candidateName = data[index][1];
  var candidateEmail = data[index][2];
  
  var date = new Date(interviewDate + 'T09:00:00');
  var event = CalendarApp.getDefaultCalendar().createEvent(
    'Interview with ' + candidateName,
    date,
    new Date(date.getTime() + 60 * 60 * 1000) // 1 hour later
  );
  
  var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "MMMM dd, yyyy 'at' hh:mm a");
  var emailBody = 'Dear ' + candidateName + ',\n\n' +
                  'We would like to invite you for an interview on ' + formattedDate + '.\n' +
                  'Please confirm if this date and time work for you.\n\n' +
                  'Best regards,\nHR Team';
  
  GmailApp.sendEmail(candidateEmail, 'Interview Invitation', emailBody);
  
  // Update the sheet with the interview date
  sheet.getRange(index + 1, 7).setValue(interviewDate); // Assuming Interview Date is in column G
}

function showScoreEntry() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  var htmlOutput = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <base target="_top">
        ${HTML_STYLES}
      </head>
      <body>
        <h2>Interview Score Entry</h2>
        <form id="scoreForm">
          <div class="form-group">
            <label for="candidate">Select Candidate:</label>
            <select id="candidate" name="candidate" required>
              <option value="">Choose a candidate</option>
  `);
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[6]) { // Only show candidates with an interview date (column G)
      htmlOutput.append(`<option value="${i}">${row[1]} - ${row[3]} (${row[6]})</option>`);
    }
  }
  
  htmlOutput.append(`
            </select>
          </div>
          <div class="form-group">
            <label for="technical">Technical Score (1-10):</label>
            <input type="number" id="technical" name="technical" min="1" max="10" required>
          </div>
          <div class="form-group">
            <label for="communication">Communication Score (1-10):</label>
            <input type="number" id="communication" name="communication" min="1" max="10" required>
          </div>
          <div class="form-group">
            <label for="cultureFit">Culture Fit Score (1-10):</label>
            <input type="number" id="cultureFit" name="cultureFit" min="1" max="10" required>
          </div>
          <input type="submit" value="Submit">
        </form>
        <script>
          document.getElementById('scoreForm').onsubmit = function(e) {
            e.preventDefault();
            var form = this;
            if (form.checkValidity()) {
              google.script.run
                .withSuccessHandler(function() {
                  form.reset();
                  alert("Score submitted successfully!");
                  google.script.host.close();
                })
                .withFailureHandler(function(error) {
                  alert("Error: " + error.message);
                })
                .addInterviewScore(
                  form.candidate.value,
                  Number(form.technical.value),
                  Number(form.communication.value),
                  Number(form.cultureFit.value)
                );
            } else {
              alert("Please fill all fields correctly.");
            }
            return false;
          };
        </script>
      </body>
    </html>
  `);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Interview Score Entry');
}

function addInterviewScore(candidateIndex, technical, communication, cultureFit) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = parseInt(candidateIndex) + 1;
  
  sheet.getRange(row, 8, 1, 3).setValues([[technical, communication, cultureFit]]); // Assuming scores start from column H
}

function generateFeedback() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var feedback = '';
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[7] && row[8] && row[9]) { // Only generate feedback for scored candidates
      var totalScore = row[7] + row[8] + row[9];
      feedback += `${row[1]} (${row[3]}): Total Score = ${totalScore}\n`;
      feedback += 'Strengths: ';
      if (row[7] >= 8) feedback += 'Technical, ';
      if (row[8] >= 8) feedback += 'Communication, ';
      if (row[9] >= 8) feedback += 'Culture Fit, ';
      feedback = feedback.slice(0, -2) + '\n';
      feedback += 'Areas of Improvement: ';
      if (row[7] <= 5) feedback += 'Technical, ';
      if (row[8] <= 5) feedback += 'Communication, ';
      if (row[9] <= 5) feedback += 'Culture Fit, ';
      feedback = feedback.slice(0, -2) + '\n\n';
    }
  }
  
  var htmlOutput = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <base target="_top">
        ${HTML_STYLES}
      </head>
      <body>
        <h2>Feedback Summary</h2>
        <pre>${feedback}</pre>
      </body>
    </html>
  `);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Feedback Summary');
}

function showDecisionSupport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  var htmlOutput = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <base target="_top">
        ${HTML_STYLES}
      </head>
      <body>
        <h2>Decision Support</h2>
        <form id="decisionForm">
          <div class="form-group">
            <label for="candidate">Select Candidate:</label>
            <select id="candidate" name="candidate" required>
              <option value="">Choose a candidate</option>
  `);
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[7] && row[8] && row[9]) { // Only show candidates with scores
      htmlOutput.append(`<option value="${i}">${row[1]} - ${row[3]}</option>`);
    }
  }
  
  htmlOutput.append(`
            </select>
          </div>
          <div class="form-group">
            <label for="decision">Decision:</label>
            <select id="decision" name="decision" required>
              <option value="">Choose a decision</option>
              <option value="Hire">Hire</option>
              <option value="No Hire">No Hire</option>
            </select>
          </div>
          <input type="submit" value="Submit">
        </form>
        <script>
          document.getElementById('decisionForm').onsubmit = function(e) {
            e.preventDefault();
            var form = this;
            if (form.checkValidity()) {
              google.script.run
                .withSuccessHandler(function() {
                  form.reset();
                  alert("Decision recorded successfully!");
                  google.script.host.close();
                })
                .withFailureHandler(function(error) {
                  alert("Error: " + error.message);
                })
                .addDecision(
                  form.candidate.value,
                  form.decision.value
                );
            } else {
              alert("Please fill all fields correctly.");
            }
            return false;
          };
        </script>
      </body>
    </html>
  `);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Decision Support');
}

function addDecision(candidateIndex, decision) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = parseInt(candidateIndex) + 1;
  
  sheet.getRange(row, 11).setValue(decision); // Assuming Decision is in column K

  // If the decision is "Hire", send the offer letter
  if (decision === 'Hire') {
    sendOfferLetter(row);
  }
}

function sendOfferLetter(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var candidateName = data[1];
  var candidateEmail = data[2];
  var position = data[3];
  
  var emailSubject = 'Job Offer for ' + position;
  var emailBody = 'Dear ' + candidateName + ',\n\n' +
                  'We are pleased to offer you the position of ' + position + ' at our company. ' +
                  'We believe your skills and experience will be a valuable asset to our team.\n\n' +
                  'Please find the attached offer letter for further details.\n\n' +
                  'Best regards,\nHR Team';
  
  var offerLetter = DriveApp.createFile('Offer_Letter_' + candidateName + '.txt', 'Congratulations ' + candidateName + '!\n\nWe are excited to offer you the position of ' + position + ' at our company.');
  
  GmailApp.sendEmail(candidateEmail, emailSubject, emailBody, {
    attachments: [offerLetter]
  });
}

function generateReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var report = 'CandidateID,CandidateName,CandidateEmail,Position,Summarization,Keyword,InterviewDate,Technical,Communication,CultureFit,Decision,Notes\n';
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    report += `${row[0]},${row[1]},${row[2]},${row[3]},${row[4]},${row[5]},${row[6]},${row[7]},${row[8]},${row[9]},${row[10]},${row[11]}\n`;
  }
  
  var blob = Utilities.newBlob(report, 'text/csv', 'Interview Report.csv');
  var file = DriveApp.createFile(blob);
  
  var htmlOutput = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <base target="_top">
        ${HTML_STYLES}
      </head>
      <body>
        <h2>Report Generated</h2>
        <p>Click <a href="${file.getUrl()}" target="_blank">here</a> to download the report.</p>
      </body>
    </html>
  `);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Report Generated');

}

function generateSampleData(numEntries) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Clear existing data
  sheet.clear();
  
  // Set headers
  sheet.getRange("A1:L1").setValues([["CandidateID", "CandidateName", "CandidateEmail", "Position", "Summarization", "Keyword", "InterviewDate", "TechnicalScore", "CommunicationScore", "CultureFitScore", "Decision", "Notes"]]);
  
  // Sample data
  var firstNames = ["John", "Jane", "Michael", "Emily", "David", "Sarah", "Robert", "Lisa", "William", "Emma"];
  var lastNames = ["Smith", "Johnson", "Brown", "Davis", "Wilson", "Moore", "Taylor", "Anderson", "Thomas", "Jackson"];
  var positions = ["Software Engineer", "Data Analyst", "Product Manager", "UX Designer", "Marketing Specialist"];
  var keywords = ["JavaScript", "Python", "SQL", "Machine Learning", "UI/UX", "Agile", "SEO", "Data Visualization"];
  var decisions = ["", "Hire", "No Hire"];
  
  // Generate random data
  var data = [];
  for (var i = 0; i < numEntries; i++) {
    var firstName = firstNames[Math.floor(Math.random() * firstNames.length)];
    var lastName = lastNames[Math.floor(Math.random() * lastNames.length)];
    var position = positions[Math.floor(Math.random() * positions.length)];
    var keyword = keywords[Math.floor(Math.random() * keywords.length)];
    
    var candidateID = "C" + (1000 + i).toString();
    var candidateName = firstName + " " + lastName;
    var candidateEmail = firstName.toLowerCase() + "." + lastName.toLowerCase() + "@example.com";
    var summarization = "Experienced " + position + " with skills in " + keyword;
    var interviewDate = new Date(2024, Math.floor(Math.random() * 12), Math.floor(Math.random() * 28) + 1);
    var technicalScore = Math.random() < 0.3 ? "" : Math.floor(Math.random() * 10) + 1;
    var communicationScore = Math.random() < 0.3 ? "" : Math.floor(Math.random() * 10) + 1;
    var cultureFitScore = Math.random() < 0.3 ? "" : Math.floor(Math.random() * 10) + 1;
    var decision = decisions[Math.floor(Math.random() * decisions.length)];
    var notes = decision ? (decision === "Hire" ? "Strong candidate" : "Not a good fit") : "";
    
    data.push([
      candidateID,
      candidateName,
      candidateEmail,
      position,
      summarization,
      keyword,
      interviewDate,
      technicalScore,
      communicationScore,
      cultureFitScore,
      decision,
      notes
    ]);
  }
  
  // Write data to sheet
  sheet.getRange(2, 1, data.length, 12).setValues(data);
  
  // Format date column
  sheet.getRange(2, 7, data.length, 1).setNumberFormat("yyyy-mm-dd");
}

// Add this to your custom menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('HR Interview')
    .addItem('List Candidates and Send Invites', 'listCandidatesAndSendInvites')
    .addItem('Enter Interview Scores', 'showScoreEntry')
    .addItem('Generate Feedback', 'generateFeedback')
    .addItem('Decision Support', 'showDecisionSupport')
    .addItem('Generate Report', 'generateReport')
    .addItem('Generate Sample Data', 'generateSampleDataPrompt')
    .addToUi();
}

function generateSampleDataPrompt() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    'Generate Sample Data',
    'How many sample entries do you want to generate?',
    ui.ButtonSet.OK_CANCEL);

  var button = result.getSelectedButton();
  var numEntries = result.getResponseText();
  
  if (button == ui.Button.OK) {
    generateSampleData(parseInt(numEntries));
    ui.alert('Sample data generated successfully!');
  }
}