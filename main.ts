// Compiled using auto-gas 1.0.0 (TypeScript 4.9.5)
// 自動化用テンプレートファイル
var DOC_TEMPLATE = DriveApp.getFileById(
  "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
);
// PDF出力先
var PDF_OUTDIR = DriveApp.getFolderById(
  "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
);
function gaslog_formSubmit(e) {
  var itemResponses;
  // フォームの回答をイベントオブジェクトまたはフォーム自身から取得する。
  if (e !== undefined) {
    itemResponses = e.response.getItemResponses();
  } else {
    var wFormRes = FormApp.getActiveForm().getResponses();
    itemResponses = wFormRes[wFormRes.length - 1].getItemResponses();
  }
  // 回答をもとに欠席届を作成する
  var wFileRtn = createGDoc(itemResponses);
}
function createGDoc(itemResponses) {
  // テンプレートファイルをコピーする
  var wCopyFile = DOC_TEMPLATE.makeCopy(),
    wCopyFileId = wCopyFile.getId(),
    wCopyDoc = DocumentApp.openById(wCopyFileId); // コピーしたファイルをGoogleドキュメントとして開く
  var wCopyDocBody = wCopyDoc.getBody(), // Googleドキュメント内の本文を取得する
    today = dayjs.dayjs();
  // 本日の日付を設定する
  wCopyDocBody = wCopyDocBody.replaceText(
    "{{date}}",
    today.format("YYYY年MM月DD日")
  );
  // 以降はGoogleフォームの回答をマッピングする
  let name = "";
  itemResponses.forEach(function (itemResponse) {
    switch (itemResponse.getItem().getTitle()) {
      case "顧客名":
        wCopyDocBody = wCopyDocBody.replaceText(
          "{{name}}",
          itemResponse.getResponse()
        );
        name = itemResponse.getResponse();
        break;
      case "作成日":
        wCopyDocBody = wCopyDocBody.replaceText(
          "{{number1}}",
          itemResponse.getResponse()
        );
        break;
      case "請求金額":
        wCopyDocBody = wCopyDocBody.replaceText(
          "{{value1}}",
          itemResponse.getResponse()
        );
        break;
      case "請求期限":
        wCopyDocBody = wCopyDocBody.replaceText(
          "{{number2}}",
          itemResponse.getResponse()
        );
        break;
      default:
        break;
    }
  });
  wCopyDoc.saveAndClose();
  // ファイル名を変更する
  let fileName = name + "様宛て請求書";
  wCopyFile.setName(fileName);
  // コピーしたファイルIDとファイル名を返却する（あとでこのIDをもとにPDFに変換するため）
  var create = createPdf(wCopyFileId, fileName);
}
function createPdf(docId, fileName) {
  // PDF変換するためのベースURLを作成する
  var wUrl = "https://docs.google.com/document/d/".concat(
    docId,
    "/export?exportFormat=pdf"
  );
  // headersにアクセストークンを格納する
  var wOtions = {
    headers: {
      Authorization: "Bearer ".concat(ScriptApp.getOAuthToken()),
    },
  };
  // PDFを作成する
  var wBlob = UrlFetchApp.fetch(wUrl, wOtions)
    .getBlob()
    .setName(fileName + ".pdf");
  //PDFを指定したフォルダに保存する
  return PDF_OUTDIR.createFile(wBlob).getId();
}
