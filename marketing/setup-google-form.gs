/**
 * BostonLuxLimos — Google Workspace native reservation intake.
 * Run setupBostonLuxLimosQuoteForm ONCE while signed in as sj@bostonluxlimos.com.
 * Creates an UNPUBLISHED Google Form and linked response Sheet in the signed-in account.
 * No Formspree, external account, password, API key or paid service required.
 * After reviewing the Form, publish it to "Anyone with the link" and enable
 * Responses > More > Get email notifications for new responses.
 *
 * The script is idempotent for this Apps Script project via Script Properties.
 * Do not copy an existing configured project and rerun without checking first.
 */
function setupBostonLuxLimosQuoteForm() {
  const props = PropertiesService.getScriptProperties();
  const prior = props.getProperty('BLL_FORM_ID');
  if (prior) {
    const existing = FormApp.openById(prior);
    Logger.log('Existing form (no duplicate created): ' + existing.getEditUrl());
    Logger.log('Existing responder URL: ' + existing.getPublishedUrl());
    Logger.log('Existing responses spreadsheet ID: ' + props.getProperty('BLL_SHEET_ID'));
    return;
  }

  const form = FormApp.create('BostonLuxLimos | Transportation Quote Request', false);

  form.setDescription(
    'Private chauffeured transportation throughout New England. ' +
    'Complete the details below to request availability and a personalized quote. ' +
    'Submitting this form is not a confirmed reservation. ' +
    'For urgent inquiries, call (212) 961-7894. ' +
    'Please do not include payment card numbers, passport details, or sensitive personal information.'
  );
  form.setConfirmationMessage(
    'Thank you for contacting BostonLuxLimos. Your quote request has been received. ' +
    'Our team will review your travel details and contact you regarding availability and pricing. ' +
    'No reservation has been confirmed. For urgent travel, call (212) 961-7894.'
  );
  form.setIsQuiz(false);
  form.setShuffleQuestions(false);
  form.setLimitOneResponsePerUser(false);
  form.setCollectEmail(false);
  form.setPublishingSummary(false);
  form.setShowLinkToRespondAgain(false);

  form.addSectionHeaderItem().setTitle('Contact information');
  form.addTextItem().setTitle('Full name').setRequired(true);
  form.addTextItem().setTitle('Email address').setRequired(true)
    .setValidation(FormApp.createTextValidation()
      .setHelpText('Please enter a valid email address.')
      .requireTextIsEmail().build());
  form.addTextItem().setTitle('Phone number').setRequired(true);

  form.addSectionHeaderItem().setTitle('Transportation details');
  form.addListItem().setTitle('Service type').setRequired(true).setChoiceValues([
    'Airport transfer', 'Corporate transportation', 'Wedding transportation',
    'Special event', 'Hourly charter', 'Long-distance transportation'
  ]);
  form.addTextItem().setTitle('Pickup location').setRequired(true);
  form.addTextItem().setTitle('Destination').setRequired(true);
  form.addDateItem().setTitle('Travel date').setRequired(true);
  form.addTimeItem().setTitle('Pickup time').setRequired(true);
  form.addTextItem().setTitle('Number of passengers').setRequired(true)
    .setValidation(FormApp.createTextValidation()
      .setHelpText('Enter a whole number from 1 to 40.')
      .requireWholeNumber().requireNumberBetween(1, 40).build());
  form.addListItem().setTitle('Journey type').setRequired(true)
    .setChoiceValues(['One way', 'Round trip', 'Hourly']);
  form.addParagraphTextItem().setTitle('Flight number, return details, luggage and special requests')
    .setHelpText('Optional. Please do not submit payment details or sensitive personal information.')
    .setRequired(false);

  const sheet = SpreadsheetApp.create('BostonLuxLimos | Reservation Quote Responses');
  props.setProperty('BLL_SHEET_ID', sheet.getId());
  form.setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId());
  props.setProperty('BLL_FORM_ID', form.getId());

  Logger.log('FORM EDITOR URL: ' + form.getEditUrl());
  Logger.log('RESPONDER URL (not accessible until published): ' + form.getPublishedUrl());
  Logger.log('RESPONSE SHEET URL: ' + sheet.getUrl());
  Logger.log('NEXT: Open the editor, verify settings, PUBLISH to Anyone with link,');
  Logger.log('enable email notifications, submit a test from an incognito browser,');
  Logger.log('then share the published responder URL for website integration.');
}
