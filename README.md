Email Generator
===============

Generates emails using our internal settings here at Frederick County.  Basically, the MIS programmer and I needed a way to make fast global changes to lots of applications in the event that IT staff change mail server settings.  We reference this class in our C# and VB projects now instead of laboriously re-writing the same code time and again.

Setup
-----

We simply add this to any project which needs to create an email.  SMTP settings are stored in a global config file and are referenced in this code.  There are two overloads of the SendEmail method: one generates an email with fewer options than the other.  This is really arbitrary and only done for our own internal ease of plugging the code in different places.  Since the SendEmail method returns a string, we can get a success message:

	// Instantiate a mail object.
	Email4.Generation mail = new Email4.Generation();
	// Send.  Set a string to the returned result.
	string strEmailResult = mail.SendEmail(strFrom, strTo, strCC, strAttachments, strSubject, strBody, strReplyTo, boolBodyHtml, strMailServer, strUser, strPass, boolSsl);
	// Display the result on the page (in this case, a literal.)
	litSuccess.Text = strEmailResult;
