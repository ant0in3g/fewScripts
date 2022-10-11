function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Emails')
    .addItem('Exam 3 Nov.', 'sendEmailsNOV')
    .addToUi();
}

function format(x, n){
    x = parseFloat(x);
    n = n || 2 ;
    return parseFloat(x.toFixed(n))
}

function sendEmailsNOV() {

  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange('A1:O101');
  var values = range.getValues();

  for (var i = 4; i < 99; i++) {

    var aliases = GmailApp.getAliases();
    var emailAddress = values[i][0];
    var subject = 'Details Note - Examen de résistance des matériaux (3 novembre 2021)';
    var message =
    "bonjour " + values[i][2] + ",\n" +
    "\n" +
    "Ta note pour l'examen de résistance des matériaux du 3 novembre est de " +  format(values[i][14]) + ".\n" +
    "La moyenne de classe est de " + format(values[100][14]) + ".\n" +
    "\n" +
    "Le détail de ta note est donné ci-dessous. Chaque question est notée sur 4. Je lui ai ensuite attribuée un certain nombres de points, donné entre parenthèses. Pour illustrer cette convention. si une question est notée sur 2 points et que tu as 3, alors cette question te rapporte 1.5 points." + "\n" +
    "\n" +
    "Exercice 1" + "\n" +
    "Question 1 " + " (" + values[3][3] + " points) : " + values[i][3] + "\n" +
    "Question 2 " + " (" + values[3][4] + " points) : " + values[i][4] + "\n" +
    "\n" +
    "Exercice 2" + "\n" +
    "Question 1 " + " (" + values[3][5] + " points) : " + values[i][5] + "\n" +
    "Question 2 " + " (" + values[3][6] + " points) : " + values[i][6] + "\n" +
    "Question 3 " + " (" + values[3][7] + " points) : " + values[i][7] + "\n" +
    "Question 4 " + " (" + values[3][8] + " points) : " + values[i][8] + "\n" +
    "Question 5 " + " (" + values[3][9] + " points) : " + values[i][9] + "\n" +
    "Question 6 " + " (" + values[3][10] + " points) : " + values[i][10] + "\n" +
    "\n" +
    "Exercice 3" + "\n" +
    "Question 1 " + " (" + values[3][11] + " points) : " + values[i][11] + "\n" +
    "Question 2 " + " (" + values[3][12] + " points) : " + values[i][12] + "\n" +
    "Question 3 " + " (" + values[3][13] + " points) : " + values[i][13] + "\n" +
    "\n" +
    "Bonne journée." + "\n" +
    "\n" +
    "Antoine Géré";

    GmailApp.sendEmail(emailAddress, subject, message, {'from': aliases[1]});

  }
}
