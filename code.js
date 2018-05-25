/**
 * @OnlyCurrentDoc
 *
 * Specifica che l'add-on potrà solamente accedere a questo documento
 * e non ad altri file.
 */

/**
 * Parametri di configurazione
 */
//var PATH = 'http://lode.disi.unitn.it';
var PATH = 'http://79.41.106.69:8080';
var REC_URL = 'http://some.url.com/recording?time=';
var EMAIL = 'giacomo.callegari@studenti.unitn.it';
var PASSWORD = 'ciao1234';
var PIN = '1234';


/**
 * Inserisce una nuova voce nel menù dei componenti aggiuntivi di Google Slides.
 * Questo metodo non è mai utilizzato nella versione mobile dell'add-on.
 *
 * @param {object} e Il parametro evento per un semplice trigger di tipo onOpen.
 *     Il tipo di autorizzazione è disponibile ispezionando e.authMode.
 */
function onOpen(e) {
    // Visualizzo la sidebar.
    SlidesApp.getUi().createAddonMenu()
        .addItem('Avvia', 'showSidebar')
        .addToUi();
}

/**
 * Viene eseguito quando l'add-on viene installato.
 * Questo metodo non è mai utilizzato nella versione mobile dell'add-on.
 *
 * @param {object} e Il parametro evento per un semplice trigger di tipo onInstall.
 *     Il tipo di autorizzazione è disponibile ispezionando e.authMode.
 */
function onInstall(e) {
    onOpen(e);
}

/**
 * Apre una sidebar nel documento contenente l'interfaccia utente dell'add-on.
 * Questo metodo non è mai utilizzato nella versione mobile dell'add-on.
 */
function showSidebar() {
    var ui = HtmlService.createHtmlOutputFromFile('sidebar')
        .setTitle('Strumento di cattura');
    SlidesApp.getUi().showSidebar(ui);
}

/**
 * Accede al sistema LODE e ottiene, se esiste, la lezione in corso.
 */
function init() {
    // Accedo al sistema LODE.
    login();

    // Ottengo l'ID della lezione.
    getLecture();
}

/**
 * Accede al sistema LODE con le credenziali richieste.
 */
function login() {
    console.log('Login');

    // Definisco l'URL della richiesta.
    var reqUrl = PATH + '/api/user/login';

    // Definisco i parametri della richiesta.
    var email = EMAIL;
    var password = PASSWORD;
    var body = {
        email: email,
        password: password
    };
    var options = {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify(body)
    };

    // Invio la richiesta.
    var response = JSON.parse(UrlFetchApp.fetch(reqUrl, options));
    console.log('Token: ' + response.token);

    // Salvo il token di autorizzazione.
    PropertiesService.getDocumentProperties().setProperty('TOKEN', response.token);
}

/**
 * Esegue una richiesta HTTP al dispositivo di cattura per ottenere la lezione in corso.
 */
function getLecture() {
    console.log('Get lecture');

    // Ottengo il token di autorizzazione.
    var token = PropertiesService.getDocumentProperties().getProperty('TOKEN');

    // Definisco l'URL della richiesta.
    var reqUrl = PATH + '/api/lecture?live=true';

    // Definisco i parametri della richiesta.
    var headers = {
        'Authorization': 'Bearer ' + token
    };
    var options = {
        method: 'GET',
        headers: headers
    };

    // Invio la richiesta.
    var response = JSON.parse(UrlFetchApp.fetch(reqUrl, options));
    console.log('Lecture: ' + response[0].uuid);

    // Salvo l'ID della lezione.
    PropertiesService.getDocumentProperties().setProperty('LEC_ID', response[0].uuid);
}

/**
 * Esegue una richiesta HTTP al dispositivo di cattura per ottenere la schermata corrente.
 * @return {Object} Il blob della schermata.
 */
function getScreenshot() {
    console.log('Get screenshot');

    // Ottengo il token di autorizzazione e l'ID della lezione.
    var token = PropertiesService.getDocumentProperties().getProperty('TOKEN');
    var lecture = PropertiesService.getDocumentProperties().getProperty('LEC_ID');

    // Definisco l'URL della richiesta.
    var reqUrl = PATH + '/api/lecture/' + lecture + '/screenshot';

    // Definisco i parametri della richiesta.
    var headers = {
        'Authorization': 'Bearer ' + token,
        'pin': PIN
    };
    var options = {
        method: 'GET',
        headers: headers
    };

    // Invio la richiesta.
    var response = JSON.parse(UrlFetchApp.fetch(reqUrl, options));
    var blob = Utilities.newBlob(Utilities.base64Decode(response.img), MimeType.PNG);

    // Restituisco il blob della schermata.
    return blob;
}

/**
 * Esegue una richiesta HTTP al dispositivo di cattura per ottenere il timestamp corrente.
 * @return {Object} Il timestamp.
 */
function getTimestamp() {
    console.log('Get timestamp');

    // Ottengo il token di autorizzazione e l'ID della lezione.
    var token = PropertiesService.getDocumentProperties().getProperty('TOKEN');
    var lecture = PropertiesService.getDocumentProperties().getProperty('LEC_ID');

    // Definisco l'URL della richiesta.
    var reqUrl = PATH + '/api/lecture/' + lecture + '/screenshot';

    // Definisco i parametri della richiesta.
    var headers = {
        'Authorization': 'Bearer ' + token,
        'pin': PIN
    };
    var options = {
        method: 'GET',
        headers: headers
    };

    // Invio la richiesta.
    var response = JSON.parse(UrlFetchApp.fetch(reqUrl, options));
    var timestamp = response.timestamp;

    //Restituisco il timestamp.
    return timestamp;
}

/**
 * Inserisce la schermata come sfondo di una nuova diapositiva e inserisce il timestamp.
 */
function insertScreenshotAndTimestamp() {
    console.log('Insert screenshot and timestamp');

    // Richiedo la schermata e il timestamp.
    var screenshot = getScreenshot();
    var timestamp = getTimestamp();

    // Inserisco una nuova diapositiva e la seleziono.
    var preso = SlidesApp.getActivePresentation();
    var newSlide = preso.appendSlide();
    newSlide.selectAsCurrentPage();

    // Creo la casella di testo per il timestamp.
    var textBox = newSlide.insertShape(SlidesApp.ShapeType.TEXT_BOX);
    var captionText = textBox.getText();
    textBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    
    //Inserisco il timestamp.
    captionText.setText('▶');
    captionText.getTextStyle().setLinkUrl(timestamp).setUnderline(false);

    // Inserisco l'immagine come sfondo della diapositiva.
    var background = newSlide.getBackground();
    background.setPictureFill(screenshot);
}

/**
 * Inserisce il timestamp nel documento.
 */
function insertTimestamp() {
    console.log('Insert timestamp');
    
    // Richiedo il timestamp.
    var timestamp = getTimestamp();

    // Inserisco una nuova diapositiva e la seleziono.
    var preso = SlidesApp.getActivePresentation();
    var newSlide = preso.appendSlide();
    newSlide.selectAsCurrentPage();

    // Creo la casella di testo per il timestamp.
    var textBox = newSlide.insertShape(SlidesApp.ShapeType.TEXT_BOX);
    var captionText = textBox.getText();
    textBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    
    //Inserisco il timestamp.
    captionText.setText('▶');
    captionText.getTextStyle().setLinkUrl(timestamp).setUnderline(false);
}
