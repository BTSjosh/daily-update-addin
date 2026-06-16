/* Southcoast Daily Update - Outlook add-in command
 * Fills the current compose email with recipients, a dated subject,
 * and the daily-update template prepended ABOVE the existing signature.
 */

Office.onReady(function () {
  // ready
});

function dailyUpdate(event) {
  var item = Office.context.mailbox.item;

  // --- Recipients (edit these to change who it goes to) ---
  item.to.setAsync([{ emailAddress: "trobinson@basesix.com" }]);
  item.cc.setAsync([
    { emailAddress: "rrowland@basesix.com" },
    { emailAddress: "chriscervantes@btsdata.com" }
  ]);

  // --- Subject with today's date as M/D ---
  var d = new Date();
  var subject = "Southcoast - Daily Update " + (d.getMonth() + 1) + "/" + d.getDate();
  item.subject.setAsync(subject);

  // --- Template body (prepended above the signature) ---
  var html =
    '<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">' +
      '<p><b>Work Completed / In Progress:</b></p>' +
      '<ul>' +
        '<li>Cameras installed: [#]</li>' +
        '<li>Data / LV cables pulled: [#]</li>' +
        '<li>Penetrations sealed: [#]</li>' +
      '</ul>' +
      '<p><b>Location(s):</b></p>' +
      '<ul>' +
        '<li>Building / Floor / Area: [ ]</li>' +
        '<li>Closet(s): [ ]</li>' +
      '</ul>' +
      '<p><b>Scope Details:</b></p>' +
      '<ul>' +
        '<li>Camera types installed: [ ]</li>' +
        '<li>Technicians on site: [ ]</li>' +
      '</ul>' +
      '<p><b>Notes / Constraints:</b></p>' +
      '<ul><li>[ ]</li></ul>' +
      '<p><b>Plan for Next Workday:</b></p>' +
      '<ul><li>[ ]</li></ul>' +
    '</div>';

  item.body.prependAsync(html, { coercionType: Office.CoercionType.Html }, function () {
    event.completed();
  });
}

// Register the function so the ribbon button can call it
Office.actions.associate("dailyUpdate", dailyUpdate);
