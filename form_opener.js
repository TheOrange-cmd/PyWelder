function updateFormSettings(formId) {
  var form = FormApp.openById(formId);
  form.setRequireLogin(false);
  form.setShowLinkToRespondAgain(true);
}