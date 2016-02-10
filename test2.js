(Saved version)

function Initialize() {

    var triggers = ScriptApp.getProjectTriggers();

    for (var i in triggers) {
        ScriptApp.deleteTrigger(triggers[i]);
    }

    ScriptApp.newTrigger("SendConfirmationMail")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();

}

function SendConfirmationMail(e) {

    try {
        //declare required variables
        var ss, columns, subject, textbody, message, senderName; //spreadsheet, columns, subject, body, htmlBody, name

        //read spreadsheet
        ss = SpreadsheetApp.getActiveSheet();
        columns = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues()[0];

        //declare variables for each field in form
        var firstName, lastName, emailAddress, tNumber, academicRank, collegeAffiliation, mentorFirstName, mentorLastName, mentorEmail, facStaffRank, multiplePresenters, ensembleName, coPresenterNames, presentationFormat, venue, equipment, title, projectType, isEDGE, projectStatus, abstract, keywords, engagementCenter, firstChoice, secondChoice;

        //define variables
        firstName = e.namedValues["First Name"].toString();
        lastName = e.namedValues["Last Name"].toString();
        emailAddress = e.namedValues["Email Address"].toString();
        tNumber = e.namedValues["T-Number"].toString();
        academicRank = e.namedValues["Academic Rank"].toString();
        collegeAffiliation = e.namedValues["College/Division Affiliation"].toString();
        mentorFirstName = e.namedValues["Mentor First Name"].toString();
        mentorLastName = e.namedValues["Mentor Last Name"].toString();
        mentorEmail = e.namedValues["Mentor Email Address"].toString();
        facStaffRank = e.namedValues["Faculty/Staff Rank"].toString();
        multiplePresenters = e.namedValues["Are there multiple presenters?"].toString();
        ensembleName = e.namedValues["Name of Ensemble"].toString();
        coPresenterNames = e.namedValues["Co-Presenter Names"].toString();
        presentationFormat = e.namedValues["Preferred Presentation Format"].toString();
        venue = e.namedValues["Preferred Venue"].toString();
        equipment = e.namedValues["Requested Equipment"].toString();
        title = e.namedValues["Title"].toString();
        projectType = e.namedValues["Project Type"].toString();
        isEDGE = e.namedValues["Is this an EDGE Project?"].toString();
        projectStatus = e.namedValues["Project Status"].toString();
        abstract = e.namedValues["Abstract/Description/Artist's Statement"].toString();
        keywords = e.namedValues["Keywords"].toString();
        engagementCenter = e.namedValues["Engagement Center"].toString();
        firstChoice = e.namedValues["First Choice"].toString();
        secondChoice = e.namedValues["Second Choice"].toString();

        // Compose email information
        sendername = "Festival of Excellence";
        subject = "Festival of Excellence - " + lastName + " - " + title;

      if(subject.length > 100){
       subject = subject.substring(0,100) + "...";
      }

        // This is the body of the auto-reply
        message = "<h1>Dear " + firstName + " " + lastName +":</h1>" +
        "<p>Thank you for registering to present <strong>\"" + title + "\"</strong> at the 2016 Festival of Excellence. Your complete submission is provided at the bottom of this email.</p>" +
        "<p>All information related to your presentation, including scheduling and room location, will be communicated via email and posted on the <a href=\"http://suu.edu/excellence\">Festival of Excellence website</a>. If applicable, please forward this and all subsequent communications to your co-presenters.</p>" +
        "<p>If you have any questions, please refer to the <a href=\"http://suu.edu/excellence/faq.html\">Frequently Asked Questions</a>. If you have further questions or concerns, please email <a href=\"mailto:festivalofexcellence@suu.edu\">festivalofexcellence@suu.edu</a>.</p>" +
        "<h2>Sincerely,<br> Festival of Excellence Committee</h2>" +
        "<p><a href=\"mailto:festivalofexcellence@suu.edu\">festivalofexcellence@suu.edu</a><br><a href=\"http://suu.edu/excellence\">suu.edu/excellence</a></p>" +
        "<hr>" +
        "<p><em>This is an automated message from SUU Festival of Excellence.</em><br>" +
        "<em>STUDENT PRESENTERS: Your mentor will also receive a copy of this email.</em></p>" +
        "<hr>" +
        "<hr>" +
        "<p>For your reference, the information you provided in your registration is below:</p>";

        // add presenter information
        message += "<h3>Presenter Information:</h3>" +
        "<p><strong>First Name</strong> | " + firstName + "<br>" +
        "<strong>Last Name</strong> | " + lastName + "<br>" +
        "<strong>Email Address</strong> | " + emailAddress + "<br>" +
        "<strong>T-Number</strong> | " + tNumber + "<br>" +
        "<strong>Academic Rank</strong> | " + academicRank + "<br>" +
        "<strong>College/Division Affiliation</strong> | " + collegeAffiliation + "</p>";

        if (mentorEmail!==""){
            // add mentor information
            message += "<h3>Mentor Information:</h3>" +
            "<p><strong>Mentor First Name</strong> | " + mentorFirstName + "<br>" +
            "<strong>Mentor Last Name</strong> | " + mentorLastName + "<br>" +
            "<strong>Mentor Email Address</strong> | " + mentorEmail + "</p>";
        }

        if (facStaffRank!==""){
            // add faculty/staff rank
            message += "<p><strong>Faculty/Staff Rank</strong> | " + facStaffRank + "</p>";
        }

        // multiple presenters
        message += "<h3>Multiple Presenters:</h3>" +
        "<p><strong>Are there multiple presenters/authors?</strong> | " + multiplePresenters + "</p>";

        if (ensembleName!==""){
            // ensemble information
            message += "<p><strong>Name of Ensemble</strong> | " + ensembleName + "</p>";
        }

        if (coPresenterNames!==""){
            // co-presenter information
            message += "<p><strong>Co-Presenter Names</strong> |<br>" +
            coPresenterNames + "</p>";
        }

        // presentation format
        message += "<h3>Presentation Format:</h3>" +
        "<p><strong>Preferred Presentation Format</strong> | " + presentationFormat + "</p>";

        if (venue!=="" || equipment!==""){
            // additional information
            message += "<p><strong>Preferred Venue</strong> | " + venue + "<br>" +
            "<strong>Requested Equipment</strong> | " + equipment + "</p>";
        }

        // presentation information
        message += "<h3>Presentation Information:</h3>" +
        "<p><strong>Title</strong> | " + title + "<br>" +
        "<strong>Project Type</strong> | " + projectType + "<br>" +
        "<strong>Is this an EDGE Project?</strong> | " + isEDGE + "<br>" +
        "<strong>Project Status</strong> | " + projectStatus + "<br>" +
        "<strong>Abstract/Description/Artist's Statement</strong> |<br>" +
        abstract + "<br>" +
        "<strong>Keywords</strong> | " + keywords + "</p>";

        if (engagementCenter!==""){
            // engagment center
            message += "<p><strong>Engagement Center</strong> | " + engagementCenter + "</p>";
        }

        // session category
        message += "<h3>Session Category:</h3>" +
        "<p><strong>First Choice</strong> | " + firstChoice + "<br>" +
        "<strong>Second Choice</strong> | " + secondChoice + "</p>";

        // strip html tags for plaintext message
        textbody = message;
        textbody = textbody.replace(/<br>/g, "\n");
        textbody = textbody.replace(/<p>/g, "");
        textbody = textbody.replace(/<\/p>/g, "\n");
        textbody = textbody.replace(/<hr>/g, "***\n");
        textbody = textbody.replace(/<h1>/g, "");
        textbody = textbody.replace(/<\/h1>/g, "\n");
        textbody = textbody.replace(/<h2>/g, "");
        textbody = textbody.replace(/<\/h2>/g, "\n");
        textbody = textbody.replace(/<h3>/g, "");
        textbody = textbody.replace(/<\/h3>/g, "\n");
        textbody = textbody.replace(/<strong>/g, "");
        textbody = textbody.replace(/<\/strong>/g, "");
        textbody = textbody.replace(/<em>/g, "");
        textbody = textbody.replace(/<\/em>/g, "");
        textbody = textbody.replace(/<p>/g, "");
        var anchor = /<a href=\\"[A-z:@\/.]*">/g;
        textbody = textbody.replace(anchor, "");
        textbody = textbody.replace(/<\/a>/g, "")

        GmailApp.sendEmail(emailAddress, subject, textbody,{
            cc: mentorEmail, name: sendername, htmlBody: message,
            });

            } catch (e) {
                Logger.log(e.toString());
            }

        }
