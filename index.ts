function onOpen() {
    const ui = DocumentApp.getUi();
    ui.createMenu('DocConverter')
        .addItem('Download Markdown', 'downloadMarkdown')
        .addItem('Download Images', 'downloadImages')
        .addItem('Download Markdown and Images', 'downloadMarkdownAndImages')
        .addItem('Download as HTML with Images', 'downloadHtmlWithImages')
        .addToUi();
}

function getFileName(exportType: string): string {
    const doc = DocumentApp.getActiveDocument();
    const title = doc.getName().replace(/[\W_]+/g, "_"); // sanitize document name
    const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
    return `${title}_${date}_${exportType}`;
}

function extractMarkdown(doc: GoogleAppsScript.Document.Document = DocumentApp.getActiveDocument()): string {
    const id = doc.getId();
    const url = `https://docs.google.com/feeds/download/documents/export/Export?exportFormat=markdown&id=${id}`;
    const response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        headers: {
            Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
        }
    });

    let content = response.getBlob().getDataAsString();

    Logger.log(content);

    // Replace image placeholders with filenames
    const body = doc.getBody();
    const images = body.getImages();
    images.forEach((image, i) => {
        const imageName = `image${i + 1}`;
        const filename = `./${imageName}.${image.getBlob().getContentType()!.split('/').pop()}`;
        Logger.log(`Replacing ${imageName} with ${filename}`);
        Logger.log(`regex: ${'!\\[.*?\\]\\[' + imageName + '\\]'}`);
        const referenceRegex = new RegExp(`!\\[.*?\\]\\[${imageName}\\]`, 'g');
        content = content.replace(referenceRegex, `![](${filename})`);
    });

    // Remove all image definitions
    Logger.log('Removing image definitions');
    Logger.log(`regex: ${'\\[.*?\\]: .*'}`);
    const imageDefinitionRegex = new RegExp(`\\[.*?\\]: .*`, 'g');
    content = content.replace(imageDefinitionRegex, '');

    Logger.log(content);
    return content;
}

function extractImages(doc: GoogleAppsScript.Document.Document = DocumentApp.getActiveDocument()): GoogleAppsScript.Base.Blob[] {
    const body = doc.getBody();
    const images = body.getImages();
    const blobs: GoogleAppsScript.Base.Blob[] = [];

    images.forEach((image, i) => {
        const blob = image.getBlob();
        blobs.push(blob.setName(`image${i + 1}.${blob.getContentType()!.split('/').pop()}`));
    });

    return blobs;
}

function downloadFile(filename: string, content: string, mimeType: string) {
    const blob = Utilities.newBlob(content, mimeType, filename);
    const html = HtmlService.createHtmlOutput(`
    <html><body onload="document.getElementById('dwn-btn').click()">
    <input type="button" id="dwn-btn" value="Download file" style="display:none;"/>
    <script>
    function download(filename, blobData) {
        const element = document.createElement('a');
        element.setAttribute('href', 'data:${mimeType};base64,' + blobData);
        element.setAttribute('download', filename);
        element.style.display = 'none';
        document.body.appendChild(element);
        element.click();
        document.body.removeChild(element);
    }
    document.getElementById("dwn-btn").addEventListener("click", function(){
        const blobData = '${Utilities.base64Encode(blob.getBytes())}';
        download("${filename}", blobData);
        window.setTimeout(function(){google.script.host.close()},100);
    }, false);
    </script>
    </body></html>  
    `).setWidth(250).setHeight(100);
    DocumentApp.getUi().showModalDialog(html, "Download file...");
}

function downloadZip(filename: string, blobs: GoogleAppsScript.Base.Blob[]) {
    const zipBlob = Utilities.zip(blobs, filename);
    const html = HtmlService.createHtmlOutput(`
    <html><body onload="document.getElementById('dwn-btn').click()">
    <input type="button" id="dwn-btn" value="Download zip file" style="display:none;"/>
    <script>
    function download(filename, blobData) {
        const element = document.createElement('a');
        element.setAttribute('href', 'data:application/zip;base64,' + blobData);
        element.setAttribute('download', filename);
        element.style.display = 'none';
        document.body.appendChild(element);
        element.click();
        document.body.removeChild(element);
    }
    document.getElementById("dwn-btn").addEventListener("click", function(){
        const blobData = '${Utilities.base64Encode(zipBlob.getBytes())}';
        download("${filename}", blobData);
        window.setTimeout(function(){google.script.host.close()},100);
    }, false);
    </script>
    </body></html>  
    `).setWidth(250).setHeight(100);
    DocumentApp.getUi().showModalDialog(html, "Download zip ...");
}

function downloadMarkdown() {
    const doc = DocumentApp.getActiveDocument();
    const content = extractMarkdown(doc);
    const filename = getFileName("markdown") + ".md";
    downloadFile(filename, content, 'text/markdown');
}

function downloadImages() {
    const doc = DocumentApp.getActiveDocument();
    const blobs = extractImages(doc);
    if (blobs.length === 0) {
        DocumentApp.getUi().alert("No images found in the document.");
        return;
    }
    const filename = getFileName("images") + ".zip";
    downloadZip(filename, blobs);
}

function downloadMarkdownAndImages() {
    const doc = DocumentApp.getActiveDocument();
    const content = extractMarkdown(doc);
    const markdownBlob = Utilities.newBlob(content, 'text/markdown', doc.getName().replace(/[\W_]+/g, "_") + ".md");
    const imageBlobs = extractImages(doc);
    const allBlobs = [markdownBlob, ...imageBlobs];
    const filename = getFileName("markdown_and_images") + ".zip";
    downloadZip(filename, allBlobs);
}

function downloadHtmlWithImages() {
    const doc = DocumentApp.getActiveDocument();
    const markdownContent = extractMarkdown(doc);

    // Use eval to load the Marked library from CDN
    eval(UrlFetchApp.fetch('https://cdn.jsdelivr.net/npm/marked/marked.min.js').getContentText());

    // Convert markdown to HTML using marked
    let htmlContent = marked.parse(markdownContent);

    const imageBlobs = extractImages(doc);
    const htmlBlob = Utilities.newBlob(htmlContent, 'text/html', doc.getName().replace(/[\W_]+/g, "_") + ".html");
    const allBlobs = [htmlBlob, ...imageBlobs];

    // Download the ZIP containing both the HTML and images
    const filename = getFileName("html_with_images") + ".zip";
    downloadZip(filename, allBlobs);
}