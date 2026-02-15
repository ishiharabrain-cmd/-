/**
 * 訪問看護連携システム - Phase 1 実装完了版
 * 
 * 主な変更点:
 * - SOAP形式プロンプトの改善（箇条書き削除、引用符削除）
 * - 在宅精神療法の自動記載（要介護2以上）
 * - 訪問時間の記録（P列、Q列）
 * - 表記修正（AI生成→クリニックで要約、患者本人_削除）
 */

var GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
var IMAGE_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('IMAGE_FOLDER_ID');
var PATIENT_SHEET_NAME = PropertiesService.getScriptProperties().getProperty('PATIENT_SHEET_NAME') || 'Patients';
var SUMMARY_FOLDER_NAME = PropertiesService.getScriptProperties().getProperty('SUMMARY_FOLDER_NAME');

// 列のインデックス（0-based）
var COL = {
  PATIENT_ID: 0,        // A: 患者番号
  NAME: 1,              // B: 患者氏名
  KANA: 2,              // C: カナ氏名
  DOB: 3,               // D: 生年月日
  GENDER: 4,            // E: 性別
  ZIP: 5,               // F: 郵便番号
  ADDRESS: 6,           // G: 住所
  ADDRESS_NOTE: 7,      // H: 住所補足
  TEL: 8,               // I: 電話番号
  TEL2: 9,              // J: 電話番号2
  FAX: 10,              // K: FAX
  EMAIL: 11,            // L: メール
  REG_DATE: 12,         // M: 登録年月日
  FACILITY: 13,         // N: 施設
  JOB: 14,              // O: 職業
  NOTE: 15,             // P: 患者備考
  TAG: 16,              // Q: タグ ★重要
  DOCTOR: 17,           // R: 主治医
  CARE_OFFICE: 18,      // S: 介護支援事業所
  CARE_MANAGER: 19,     // T: ケアマネージャー
  DEATH_DATE: 20,       // U: 死亡年月日
  DEATH_PLACE: 21,      // V: 死亡場所
  MITORIRELATION: 22,          // W: 看取り
  START_DATE: 23,       // X: 診療開始日
  END_DATE: 24,         // Y: 診療終了日
  END_REASON: 25,       // Z: 診療終了理由
  CARE_LEVEL: 26,       // AA: 要介護度 ★重要
  OPINION: 27,          // AB: 意見書当院
  HOKAN_ST: 28,         // AC: 訪看ST名
  SHIJI_TYPE: 29,       // AD: 指示書の種別
  PHARMACY: 30,         // AE: 薬局
  LAST_VISIT: 31,       // AF: 最終診療日
  STATUS: 32,           // AG: 状態
  PAYMENT: 33,          // AH: 入金方法
  NOTE2: 34             // AI: 患者備考２
};

/**
 * Webアプリのエントリーポイント
 */
function doGet(e) {
  try {
    var template = HtmlService.createTemplateFromFile('LoginPage_Real');
    return template.evaluate()
      .setTitle('訪問看護連携システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
  } catch (error) {
    return HtmlService.createHtmlOutput("<h2>エラー</h2><p>" + error.toString() + "</p>");
  }
}

/**
 * タグ一覧を取得
 */
/**
 * 患者データシートを取得（柔軟な名前検出）
 */
function getPatientSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // スクリプトプロパティから取得
  if (PATIENT_SHEET_NAME) {
    var sheet = ss.getSheetByName(PATIENT_SHEET_NAME);
    if (sheet) return sheet;
  }
  
  // 候補を順番に試す
  var candidates = ['Patients', 'kanja_list', 'Sheet1', 'シート1'];
  for (var i = 0; i < candidates.length; i++) {
    var sheet = ss.getSheetByName(candidates[i]);
    if (sheet) return sheet;
  }
  
  // 最初のシートを返す
  return ss.getSheets()[0];
}

function getTagList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Tags");
    if (!sheet) return { success: false, msg: "Tagsシートが見つかりません" };
    
    var data = sheet.getDataRange().getValues();
    var tags = [];
    
    for (var i = 1; i < data.length; i++) {
      tags.push({
        name: data[i][0],
        description: data[i][2] || ""
      });
    }
    
    return { success: true, tags: tags };
  } catch (error) {
    return { success: false, msg: error.toString() };
  }
}

/**
 * タグで認証
 */
function authenticateTag(tagName, password) {
  try {
    if (!tagName || !password) {
      return { success: false, msg: "タグ名とパスワードを入力してください" };
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Tags");
    if (!sheet) return { success: false, msg: "Tagsシートが見つかりません" };
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(tagName).trim()) {
        var storedHash = data[i][1];
        var inputHash = Utilities.base64Encode(
          Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password)
        );
        
        if (storedHash === inputHash) {
          return { 
            success: true, 
            tag: tagName,
            description: data[i][2] || ""
          };
        } else {
          return { success: false, msg: "パスワードが正しくありません" };
        }
      }
    }
    
    return { success: false, msg: "タグが見つかりません" };
  } catch (error) {
    return { success: false, msg: "エラー: " + error.toString() };
  }
}

/**
 * 患者ID+生まれた年で認証
 */
function authenticatePatient(patientId, birthYear) {
  try {
    if (!patientId || !birthYear) {
      return { success: false, msg: "患者番号と認証コードを入力してください" };
    }
    
    // 4桁の数字形式チェック（YYYY）
    if (!/^\d{4}$/.test(birthYear)) {
      return { success: false, msg: "認証コードは4桁の数字で入力してください" };
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var pSheet = getPatientSheet_();
    
    if (!pSheet) {
      return { success: false, msg: "患者データシートが見つかりません" };
    }
    
    var data = pSheet.getDataRange().getValues();
    
    // 入力された4桁の年
    var inputYear = parseInt(birthYear);
    
    Logger.log("🔍 患者認証: ID=" + patientId + ", 認証コード入力あり");
    
    for (var i = 1; i < data.length; i++) {
      var rowPatientId = String(data[i][COL.PATIENT_ID]).trim();
      
      if (rowPatientId === String(patientId).trim()) {
        var rowBirthdate = data[i][COL.DOB]; // D列（正しい定義を使用）
        
        Logger.log("🔍 D列の値: " + rowBirthdate + " (型: " + typeof rowBirthdate + ")");
        
        // 生年（年のみ）の比較
        var match = false;
        var actualYear = null;
        
        if (rowBirthdate instanceof Date) {
          // Date型の場合
          actualYear = rowBirthdate.getFullYear();
          Logger.log("📅 Date型: 抽出年=" + actualYear + ", 入力年=" + inputYear);
          match = (actualYear === inputYear);
        } else if (typeof rowBirthdate === 'string' || typeof rowBirthdate === 'number') {
          // 文字列または数値の場合
          var bdStr = String(rowBirthdate);
          Logger.log("📝 文字列: " + bdStr);
          
          // パターン1: "1950/01/15" または "1950-01-15"
          if (bdStr.indexOf('/') !== -1 || bdStr.indexOf('-') !== -1) {
            var bdParts = bdStr.replace(/-/g, '/').split('/');
            if (bdParts.length >= 3) {
              actualYear = parseInt(bdParts[0]);
              Logger.log("📅 スラッシュ形式: 抽出年=" + actualYear + ", 入力年=" + inputYear);
              match = (actualYear === inputYear);
            }
          }
          // パターン2: "19500115" (8桁数字)
          else if (/^\d{8}$/.test(bdStr)) {
            actualYear = parseInt(bdStr.substring(0, 4));
            Logger.log("📅 8桁形式: 抽出年=" + actualYear + ", 入力年=" + inputYear);
            match = (actualYear === inputYear);
          }
          // パターン3: "1950" (4桁数字、年のみ)
          else if (/^\d{4}$/.test(bdStr)) {
            actualYear = parseInt(bdStr);
            Logger.log("📅 4桁形式: 抽出年=" + actualYear + ", 入力年=" + inputYear);
            match = (actualYear === inputYear);
          }
        } else {
          Logger.log("⚠️ 想定外のデータ型: " + typeof rowBirthdate);
        }
        
        if (match) {
          Logger.log("✅ 認証成功: 患者ID=" + rowPatientId + ", 抽出年=" + actualYear + ", 入力年=" + inputYear);
          
          // Date型を文字列に変換
          var birthdateStr = rowBirthdate instanceof Date 
            ? Utilities.formatDate(rowBirthdate, Session.getScriptTimeZone(), "yyyy/MM/dd")
            : String(rowBirthdate || "");
          
          return {
            success: true,
            patient: {
              id: rowPatientId,
              name: data[i][COL.NAME],
              birthdate: birthdateStr,
              address: data[i][COL.ADDRESS] || "",
              phone: data[i][COL.TEL] || "",
              doctor: data[i][COL.DOCTOR] || "",
              pharmacy: data[i][COL.PHARMACY] || "",
              careLevel: data[i][COL.CARE_LEVEL] || "",
              note: data[i][COL.NOTE] || ""
            }
          };
        } else {
          Logger.log("❌ 認証コード不一致: 抽出年=" + actualYear + ", 入力年=" + inputYear);
          return { success: false, msg: "認証コードが一致しません" };
        }
      }
    }
    
    Logger.log("❌ 患者番号が見つかりません: " + patientId);
    return { success: false, msg: "患者番号または認証コードが正しくありません" };
  } catch (error) {
    Logger.log("❌ 患者認証エラー: " + error.toString());
    return { success: false, msg: "エラー: " + error.toString() };
  }
}

/**
 * タグに紐づく患者リストを取得
 */
function getPatientsByTag(tagName) {
  try {
    var sheet = getPatientSheet_();
    if (!sheet) return { success: false, msg: "患者データシートが見つかりません" };
    
    var data = sheet.getDataRange().getValues();
    var patients = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var tags = String(row[COL.TAG] || "").split(/[,，、]/); // カンマ（, ， 、）で分割
      var hasTag = false;
      
      for (var j = 0; j < tags.length; j++) {
        if (tags[j].trim() === tagName) {
          hasTag = true;
          break;
        }
      }
      
      if (hasTag) {
        patients.push({
          id: row[COL.PATIENT_ID],
          name: row[COL.NAME],
          kana: row[COL.KANA] || "",
          dob: row[COL.DOB] instanceof Date ? Utilities.formatDate(row[COL.DOB], Session.getScriptTimeZone(), "yyyy/MM/dd") : "",
          address: row[COL.ADDRESS] || "",
          careLevel: row[COL.CARE_LEVEL] || "",
          hokanSt: row[COL.HOKAN_ST] || "",
          status: row[COL.STATUS] || "",
          note: row[COL.NOTE] || ""
        });
      }
    }
    
    return { success: true, patients: patients };
  } catch (error) {
    return { success: false, msg: "エラー: " + error.toString() };
  }
}

/**
 * 権限承認用：強制的に外部APIアクセスをリクエスト
 * この関数を実行すると、必ず権限確認ダイアログが表示されます
 */
function forceAuthorization() {
  // 外部APIへのアクセスを試みる（権限確認のため）
  try {
    var url = "https://www.google.com";
    var response = UrlFetchApp.fetch(url);
    Logger.log("✅ 外部API呼び出し権限が承認されています");
    Logger.log("✅ これでGemini APIも使えます！");
  } catch (e) {
    Logger.log("❌ まだ権限が承認されていません");
    Logger.log("このエラーが出る場合は、承認ダイアログで「許可」をクリックしてください");
  }
}

/**
 * テスト用：AI要約機能のテスト実行
 * Apps Scriptエディタで実行→この関数を選択して実行
 */
function testAISummary() {
  Logger.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  Logger.log("🧪 AI要約機能のテスト実行");
  Logger.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  
  // 1. APIキーの確認
  if (!GEMINI_API_KEY) {
    Logger.log("❌ エラー: GEMINI_API_KEYが設定されていません");
    Logger.log("💡 対処法: スプレッドシートのメニュー「🏥 訪問看護連携」→「🔑 APIキーを設定」");
    return;
  }
  Logger.log("✅ APIキーが設定されています");
  
  // 2. 簡単なテストプロンプト
  var testPrompt = "こんにちは。テストです。";
  Logger.log("📝 テストプロンプト: " + testPrompt);
  
  // 3. API呼び出し
  var result = callGeminiAPI_(testPrompt);
  
  // 4. 結果表示
  Logger.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  Logger.log("📊 テスト結果:");
  Logger.log(result);
  Logger.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  
  if (result.indexOf("（要約は後ほど更新されます）") > -1) {
    Logger.log("❌ エラーが発生しています。上記のログを確認してください。");
  } else {
    Logger.log("✅ テスト成功！Gemini APIが正常に動作しています。");
    Logger.log("💡 次は実際の報告を送信してテストしてください。");
  }
}

/**
 * 患者の過去の記録を取得（最新5件）
 */
function getPatientHistory(patientId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var reportSheet = ss.getSheetByName("Reports");
    
    if (!reportSheet) {
      return { success: false, msg: "Reportsシートが見つかりません" };
    }
    
    var data = reportSheet.getDataRange().getValues();
    var history = [];
    
    // 患者シートから患者名を取得
    var pSheet = getPatientSheet_();
    var patientName = "";
    if (pSheet) {
      var patientData = pSheet.getDataRange().getValues();
      for (var i = 1; i < patientData.length; i++) {
        if (String(patientData[i][COL.PATIENT_ID]).trim() === String(patientId).trim()) {
          patientName = patientData[i][COL.NAME];
          break;
        }
      }
    }
    
    // Googleドキュメント・PDFのサマリーを取得
    var docSummary = getDocumentSummary_(patientId, patientName);
    
    // サマリーを先頭に追加（複数ファイルの場合は統合表示）
    if (docSummary.found) {
      var fileInfo = docSummary.files && docSummary.files.length > 0 
        ? docSummary.files.join(", ") 
        : "サマリー";
      
      history.push({
        date: "📄 " + fileInfo + " (" + docSummary.files.length + "件)",
        scenario: "カルテサマリー",
        content: docSummary.content,
        aiSummary: docSummary.summary,
        isDocSummary: true
      });
    }
    
    // この患者の全報告を取得
    var reports = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[2]).trim() === String(patientId).trim()) { // C列: 患者ID
        // 訪問開始時間（P列, index 15）を優先し、なければ登録日時（B列, index 1）を使う
        var visitDate = null;
        var visitDateDisplay = "";
        
        if (row[15]) {
          // startTime がある場合
          try {
            visitDate = new Date(row[15]);
            if (isNaN(visitDate.getTime())) {
              visitDate = row[1] instanceof Date ? row[1] : new Date(row[1]);
            }
          } catch (e) {
            visitDate = row[1] instanceof Date ? row[1] : new Date(row[1]);
          }
        } else {
          visitDate = row[1] instanceof Date ? row[1] : new Date(row[1]);
        }
        
        visitDateDisplay = visitDate instanceof Date && !isNaN(visitDate.getTime())
          ? Utilities.formatDate(visitDate, Session.getScriptTimeZone(), "MM/dd HH:mm")
          : String(row[1]);
        
        reports.push({
          reportId: row[0],  // A列: UUID
          date: visitDateDisplay,
          visitDate: visitDate,  // ソート用
          scenario: row[6] || "",
          content: row[7] || "",
          aiSummary: row[11] || "",
          reporter: row[4] || "", // E列: 記録者
          manualEdit: row[17] || false,
          manualEditDate: row[18] || "",
          manualEditBy: row[19] || "" // T列: 修正者名（新規追加列）
        });
      }
    }
    
    // 訪問日時の新しい順にソート
    reports.sort(function(a, b) {
      var dateA = a.visitDate instanceof Date ? a.visitDate.getTime() : 0;
      var dateB = b.visitDate instanceof Date ? b.visitDate.getTime() : 0;
      return dateB - dateA; // 新しい順
    });
    
    // ソート済みの報告をhistoryに追加
    for (var j = 0; j < reports.length; j++) {
      history.push({
        reportId: reports[j].reportId,
        date: reports[j].date,
        scenario: reports[j].scenario,
        content: reports[j].content,
        aiSummary: reports[j].aiSummary,
        reporter: reports[j].reporter,
        manualEdit: reports[j].manualEdit,
        manualEditDate: reports[j].manualEditDate,
        manualEditBy: reports[j].manualEditBy
      });
    }
    
    Logger.log("📚 取得した履歴数: " + history.length + "件（ドキュメント含む、訪問日時順）");
    
    return {
      success: true,
      history: history,
      totalCount: reports.length
    };
    
  } catch (e) {
    Logger.log('getPatientHistory error: ' + e.toString());
    return { success: false, msg: "記録の取得に失敗しました: " + e.toString() };
  }
}

/**
 * 患者詳細を取得
 */
function getPatientDetail(patientId, tagName) {
  try {
    var sheet = getPatientSheet_();
    if (!sheet) return { success: false, msg: "患者データシートが見つかりません" };
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[COL.PATIENT_ID]).trim() === String(patientId).trim()) {
        // タグの権限チェック（患者本人の場合はスキップ）
        if (tagName !== "患者本人") {
          var tags = String(row[COL.TAG] || "").split(/[,，、]/); // カンマ（, ， 、）で分割
          var hasAccess = false;
          
          for (var j = 0; j < tags.length; j++) {
            if (tags[j].trim() === tagName) {
              hasAccess = true;
              break;
            }
          }
          
          if (!hasAccess) {
            return { success: false, msg: "この患者にアクセスする権限がありません" };
          }
        }
        
        return {
          success: true,
          patient: {
            id: row[COL.PATIENT_ID],
            name: row[COL.NAME],
            kana: row[COL.KANA] || "",
            dob: row[COL.DOB] instanceof Date ? Utilities.formatDate(row[COL.DOB], Session.getScriptTimeZone(), "yyyy/MM/dd") : "",
            gender: row[COL.GENDER] || "",
            zip: row[COL.ZIP] || "",
            address: row[COL.ADDRESS] || "",
            addressNote: row[COL.ADDRESS_NOTE] || "",
            tel: row[COL.TEL] || "",
            tel2: row[COL.TEL2] || "",
            email: row[COL.EMAIL] || "",
            facility: row[COL.FACILITY] || "",
            note: row[COL.NOTE] || "",
            tags: row[COL.TAG] || "",
            doctor: row[COL.DOCTOR] || "",
            careOffice: row[COL.CARE_OFFICE] || "",
            careManager: row[COL.CARE_MANAGER] || "",
            careLevel: row[COL.CARE_LEVEL] || "",
            hokanSt: row[COL.HOKAN_ST] || "",
            shijiType: row[COL.SHIJI_TYPE] || "",
            pharmacy: row[COL.PHARMACY] || "",
            lastVisit: row[COL.LAST_VISIT] instanceof Date ? Utilities.formatDate(row[COL.LAST_VISIT], Session.getScriptTimeZone(), "yyyy/MM/dd") : "",
            status: row[COL.STATUS] || "",
            payment: row[COL.PAYMENT] || "",
            note2: row[COL.NOTE2] || ""
          }
        };
      }
    }
    
    return { success: false, msg: "患者が見つかりません" };
  } catch (error) {
    return { success: false, msg: "エラー: " + error.toString() };
  }
}

/**
 * 報告送信（画像URL付き）
 */
function sendReportWithImages(patientId, patientName, tagName, reporter, scenario, content, imageUrls, voiceText) {
  try {
    if (!patientId || !content) {
      return { success: false, msg: "患者番号または報告内容が空です" };
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Reports");
    if (!sheet) return { success: false, msg: "Reportsシートが見つかりません" };
    
    var imageUrlsStr = Array.isArray(imageUrls) ? imageUrls.join(', ') : "";
    
    sheet.appendRow([
      Utilities.getUuid(),
      new Date(),
      String(patientId),
      patientName || "",
      reporter || "ゲスト",
      tagName || "",
      scenario || "",
      String(content),
      imageUrlsStr,
      voiceText || "",
      false,
      "",
      ""
    ]);
    
    // 要約生成を実行
    try {
      Logger.log("📤 報告送信完了。要約生成を開始します。");
      processLatestReport_();
    } catch (e) {
      Logger.log("❌ 要約生成でエラーが発生しました: " + e.toString());
      Logger.log("スタックトレース: " + e.stack);
    }
    
    // 患者情報のサマリーを自動更新
    try {
      updatePatientSummary_(patientId);
    } catch (e) {
      Logger.log("⚠️ 患者サマリー更新でエラー: " + e.toString());
    }
    
    return { success: true };
  } catch (error) {
    return { success: false, msg: "エラー: " + error.toString() };
  }
}

/**
 * 画像をGoogle Driveにアップロード
 */
function uploadImageToDrive(base64Data, fileName) {
  try {
    if (!IMAGE_FOLDER_ID) {
      return { success: false, msg: "画像フォルダが設定されていません" };
    }
    
    var base64Clean = base64Data.replace(/^data:image\/\w+;base64,/, '');
    var blob = Utilities.newBlob(
      Utilities.base64Decode(base64Clean),
      'image/jpeg',
      fileName || 'image_' + new Date().getTime() + '.jpg'
    );
    
    var folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return { 
      success: true, 
      url: file.getUrl(),
      id: file.getId()
    };
  } catch (error) {
    return { success: false, msg: "アップロードエラー: " + error.toString() };
  }
}

/**
 * 音声ファイルをGoogle Driveにアップロードし、文字起こしを実行
 */
function uploadAudioToDrive(base64Data, fileName) {
  try {
    if (!IMAGE_FOLDER_ID) {
      return { success: false, msg: "画像フォルダが設定されていません" };
    }
    
    // audio/mpeg, audio/mp4, audio/wav などに対応
    var mimeType = 'audio/mpeg';
    if (base64Data.indexOf('data:audio/') === 0) {
      var mimeMatch = base64Data.match(/data:(audio\/[^;]+);/);
      if (mimeMatch) mimeType = mimeMatch[1];
    }
    
    var base64Clean = base64Data.replace(/^data:audio\/\w+;base64,/, '');
    var blob = Utilities.newBlob(
      Utilities.base64Decode(base64Clean),
      mimeType,
      fileName || 'audio_' + new Date().getTime() + '.mp3'
    );
    
    var folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // 音声ファイルの文字起こしを実行
    var transcription = "";
    if (GEMINI_API_KEY) {
      Logger.log("🎤 音声ファイルの文字起こしを開始...");
      transcription = transcribeAudioWithGemini_(base64Clean, mimeType);
    }
    
    return { 
      success: true, 
      url: file.getUrl(),
      id: file.getId(),
      transcription: transcription
    };
  } catch (error) {
    Logger.log("音声アップロードエラー: " + error.toString());
    return { success: false, msg: "アップロードエラー: " + error.toString() };
  }
}

/**
 * 報告を送信（音声ファイル・時刻対応版）
 */
function sendReportWithAudio(patientId, patientName, tagName, reporter, scenario, content, imageUrls, voiceText, audioUrl, reporterRole, isUrgent, startTime, endTime) {
  try {
    if (!patientId) {
      return { success: false, msg: "患者番号が空です" };
    }
    
    if (!content && (!imageUrls || imageUrls.length === 0) && !audioUrl) {
      return { success: false, msg: "報告内容、画像、または音声ファイルのいずれかを入力してください" };
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Reports");
    if (!sheet) return { success: false, msg: "Reportsシートが見つかりません" };
    
    var imageUrlsStr = Array.isArray(imageUrls) ? imageUrls.join(', ') : "";
    
    var rolePrefix = reporterRole === 'doctor' ? '[医師]' : '[看護]';
    var reporterWithRole = reporter + rolePrefix;
    
    var startTimeStr = startTime || "";
    var endTimeStr = endTime || "";
    
    sheet.appendRow([
      Utilities.getUuid(),
      new Date(),
      String(patientId),
      patientName || "",
      reporterWithRole,
      tagName || "",
      scenario || "",
      String(content || ""),
      imageUrlsStr,
      voiceText || "",
      false,
      "",
      "",
      audioUrl || "",
      reporterRole || "nurse",
      startTimeStr,
      endTimeStr
    ]);
    
    if (reporterRole === 'doctor' && isUrgent === true && content) {
      Logger.log("📋 医師リマインドメールを送信します");
      sendDoctorReminderEmail(patientId, patientName, reporter, content);
    }
    
    try {
      Logger.log("📤 報告送信完了。要約生成を開始します。");
      processLatestReport_(startTimeStr, endTimeStr);
    } catch (e) {
      Logger.log("❌ 要約生成でエラーが発生しました: " + e.toString());
    }
    
    try {
      updatePatientSummary_(patientId);
    } catch (e) {
      Logger.log("⚠️ 患者サマリー更新でエラー: " + e.toString());
    }
    
    return { success: true };
  } catch (error) {
    return { success: false, msg: "エラー: " + error.toString() };
  }
}

/**
 * 患者情報のサマリーを自動更新（過去の報告から生成）
 */
function updatePatientSummary_(patientId) {
  Logger.log("📊 患者サマリーの更新開始: " + patientId);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pSheet = getPatientSheet_();
  var rSheet = ss.getSheetByName("Reports");
  
  if (!pSheet || !rSheet) {
    Logger.log("⚠️ シートが見つかりません");
    return;
  }
  
  // この患者の過去10件の報告を取得
  var reportData = rSheet.getDataRange().getValues();
  var recentReports = [];
  
  for (var i = reportData.length - 1; i > 0; i--) {
    if (String(reportData[i][2]).trim() === String(patientId).trim()) {
      var report = {
        date: reportData[i][1],
        scenario: reportData[i][6],
        content: reportData[i][7],
        aiSummary: reportData[i][11],
        reporter: reportData[i][4] || "",
        manualEdit: reportData[i][17] || false
      };
      recentReports.push(report);
      
      if (recentReports.length >= 10) break;
    }
  }
  
  if (recentReports.length === 0) {
    Logger.log("ℹ️ 報告が見つかりません");
    return;
  }
  
  if (GEMINI_API_KEY) {
    var prompt = "あなたは訪問診療のベテラン看護師です。以下の直近の報告履歴を読み取り、" +
      "他のスタッフが見て分かりやすい自然な経過まとめを作成してください。\n\n" +
      "指示事項:\n" +
      "1. 硬い項目分けはせず、2〜3文の自然な文章をつなげて記述\n" +
      "2. 口調は「〜です」「〜されています」「〜とのことです」といった丁寧な申し送り調\n" +
      "3. 箇条書きを多用せず、特筆すべき変化や注意が必要なポイントに絞る\n" +
      "4. AIが作成したことを感じさせる前置きや結びは一切不要\n" +
      "5. 箇条書きの記号（・、-、*）は使用しない\n" +
      "6. 引用符（\"、『』、「」）は使用しない\n" +
      "7. 手動修正された記録は、その内容を尊重して優先的に反映する\n\n" +
      "最近の報告内容:\n";
    
    recentReports.forEach(function(report, index) {
      prompt += "---\n";
      prompt += "日時: " + report.date + "\n";
      if (report.manualEdit) {
        prompt += "【※手動修正済み - この内容を優先的に参考にする】\n";
      }
      prompt += "内容: " + (report.aiSummary || report.content) + "\n";
    });
    
    Logger.log("🤖 患者サマリー（備考2）を自然な文章で生成中...");
    var patientSummary = callGeminiAPI_(prompt);
    
    // 患者名を取得
    var patientName = "";
    var patientData = pSheet.getDataRange().getValues();
    for (var i = 1; i < patientData.length; i++) {
      if (String(patientData[i][COL.PATIENT_ID]).trim() === String(patientId).trim()) {
        patientName = patientData[i][COL.NAME];
        pSheet.getRange(i + 1, COL.NOTE2 + 1).setValue(patientSummary);
        Logger.log("✅ 患者サマリーを更新しました（行" + (i + 1) + "）");
        break;
      }
    }
    
    // Summaryシートにも記録を保存（最近の経過として）
    var summarySheet = ss.getSheetByName("Summary");
    if (summarySheet) {
      var existingRecord = getSummaryRecord_(summarySheet, patientId);
      
      // 最近の経過ログを構築
      var updateLog = buildUpdateLog_(recentReports);
      
      if (existingRecord) {
        // 既存レコードを更新
        summarySheet.getRange(existingRecord.row, 12).setValue(updateLog);
        summarySheet.getRange(existingRecord.row, 13).setValue((existingRecord.updateCount || 0) + 1);
      } else {
        // 新規レコードを作成
        saveSummaryRecord_(summarySheet, {
          patientId: patientId,
          patientName: patientName,
          summary: "",
          sourceFiles: "",
          fileCount: 0,
          recentUpdates: updateLog,
          updateCount: 1
        });
      }
      Logger.log("✅ Summaryシートも更新しました");
    }
  }
}

/**
 * 最新の未処理報告をAI処理
 */
function processLatestReport_(startTime, endTime) {
  Logger.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  Logger.log("🤖 AI処理開始");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reportSheet = ss.getSheetByName("Reports");
  
  if (!reportSheet) {
    Logger.log("❌ Reportsシートが見つかりません");
    return;
  }
  
  var data = reportSheet.getDataRange().getValues();
  Logger.log("📊 データ行数: " + (data.length - 1));
  
  for (var i = data.length - 1; i >= 1; i--) {
    var row = data[i];
    var processed = row[10];
    
    Logger.log("🔍 行" + (i + 1) + "を確認: AI_Processed = " + processed);
    
    if (processed === false || processed === "") {
      var patientId = row[2];
      var content = row[7];
      var imageUrls = row[8];
      var reporterRole = row[14] || "nurse";
      
      Logger.log("📝 未処理の報告を発見");
      Logger.log("  患者ID: " + patientId);
      Logger.log("  役割: " + reporterRole);
      
      if (!GEMINI_API_KEY) {
        Logger.log("⚠️ APIキーが未設定のため、要約生成をスキップ");
        reportSheet.getRange(i + 1, 11).setValue(true);
        reportSheet.getRange(i + 1, 12).setValue("（要約は後ほど更新されます）");
        reportSheet.getRange(i + 1, 13).setValue(generateChartText_(content, patientId, row[3], imageUrls));
      } else {
        Logger.log("🤖 AI要約を生成中...");
        
        var reporterName = row[4] || ""; // 記録者名を取得
        var patientName = row[3] || ""; // 患者名を取得
        var summary = "";
        if (reporterRole === 'doctor') {
          Logger.log("👨‍⚕️ 医師モード: SOAP形式で要約生成");
          summary = generateSOAPSummary_(content, patientId, startTime, endTime, reporterName);
        } else {
          Logger.log("👩‍⚕️ 看護師モード: 申し送り形式で要約生成");
          summary = generateAISummary_(content, reporterName, patientName, startTime, endTime);
          sendDoctorAlertIfRequested_(summary, patientId, row[3], row[4]);
        }
        
        Logger.log("📋 カルテテキストを生成中...");
        var chartText = generateChartText_(content, patientId, row[3], imageUrls);
        
        Logger.log("💾 結果を保存中...");
        reportSheet.getRange(i + 1, 11).setValue(true);
        reportSheet.getRange(i + 1, 12).setValue(summary);
        reportSheet.getRange(i + 1, 13).setValue(chartText);
        Logger.log("✅ AI処理完了");
      }
      
      break;
    }
  }
  
  Logger.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
}

/**
 * 看護記録形式の要約生成 - Phase 1 改善版
 */
function generateAISummary_(content, reporterName, patientName, startTime, endTime) {
  if (!GEMINI_API_KEY) return "";
  
  // ヘッダー情報を構築（患者名・訪問時間・記録者）
  var headerText = "";
  if (patientName) {
    headerText += "患者名: " + patientName + " 様\n\n";
  }
  
  // 訪問時間の整形
  var visitTimeText = "";
  if (startTime && endTime) {
    try {
      var start = new Date(startTime);
      var end = new Date(endTime);
      var duration = Math.round((end - start) / 60000);
      
      var startStr = formatTime_(start);
      var endStr = formatTime_(end);
      
      visitTimeText = "訪問時間: " + startStr + " - " + endStr + "\n";
      if (duration > 0) {
        visitTimeText += "訪問時間: " + duration + "分\n";
      }
      visitTimeText += "\n";
    } catch (e) {
      Logger.log("⚠️ 時間情報の整形エラー: " + e.toString());
    }
  }
  
  headerText += visitTimeText;
  
  var reporterInfo = reporterName ? "記録者: " + reporterName + "\n\n" : "";
  
  var prompt = "あなたは経験豊富な訪問看護師です。以下の報告を読み、他のスタッフが見て分かりやすい申し送りを作成してください。\n\n" +
    reporterInfo +
    "報告内容:\n" + content + "\n\n" +
    "重要：報告内容の冒頭に【確認事項】がある場合は、必ずその内容（診療種別、処方せん変更の有無、通院／在宅精神療法の有無）を申し送りの最後に含めてください。\n\n" +
    "出力の冒頭には、以下のヘッダー情報をそのまま記載してください:\n" +
    headerText + "\n" +
    "出力形式（項目名は【 】で囲み、各項目は改行で区切る）:\n\n" +
    "【バイタルサイン】\n" +
    "報告にバイタルデータがある場合、体温、血圧、脈拍、SPO2、呼吸数などの数値を一字一句変えずにそのまま記載し、その安定度や変化を一文で添える。\n" +
    "例: 体温36.3℃、血圧120/80mmHg、脈拍78回(整)、呼吸数20回、SPO2 98%で安定しています。\n\n" +
    "【本日の状態】\n" +
    "患者の主訴や様子を2〜3文で自然に記述する。\n\n" +
    "【栄養・水分】\n" +
    "飲水量、食事内容など栄養関連の情報がある場合に記載する。数値はそのまま記載する。\n" +
    "例: 飲水量は450ml。\n\n" +
    "【排泄】\n" +
    "排便・排尿に関する情報がある場合に記載する。日付はそのまま記載する。\n" +
    "例: 最終排便は2月10日。明日グリセリン浣腸（GE）と摘便を実施する予定です。\n\n" +
    "【ケア実施内容】\n" +
    "実施したケアがある場合に記載する。\n" +
    "例: オムツ交換、陰部洗浄、軟膏塗布を実施しています。\n\n" +
    "【特記事項】\n" +
    "上記に含まれない注意事項があれば記載する。なければこのセクション自体を省略する。\n\n" +
    "【医師への依頼】\n" +
    "処方変更・検査依頼などがあれば明記する。なければこのセクション自体を省略する。\n\n" +
    "【次回訪問】\n" +
    "次回訪問日時がある場合はそのまま記載し、確認すべき重点項目を添える。\n" +
    "例: 訪看は2月14日(土) 14時30分頃。排便の有無を確認し、排便処置を実施してください。\n\n" +
    "【確認事項】（医師報告の場合のみ）\n" +
    "記録者: （記録者名）\n" +
    "診療種別: （訪問/外来）\n" +
    "処方せん変更: （あり/なし）\n" +
    "通院／在宅精神療法: （あり/なし）\n" +
    "※報告内容の冒頭に【確認事項】がある場合のみ出力し、ない場合は省略する。\n\n" +
    "最重要ルール（数値の正確な転記）:\n" +
    "報告内容に含まれる数値データ（体温、血圧、脈拍、SPO2、呼吸数、飲水量、日付、時刻など）は絶対に変更・丸め・省略しない。報告にある数値をそのまま転記すること。\n" +
    "例: 報告に「36.3℃」とあれば「36.3℃」と書く。「安定」や「正常範囲」のように数値を言い換えてはならない。\n\n" +
    "厳守事項:\n" +
    "- Markdown記法（見出し記号#、太字**）は一切使用禁止\n" +
    "- 箇条書きの記号（・、-、*、ハイフン）は使用しない\n" +
    "- 項目名は必ず【 】で囲む\n" +
    "- 各項目は改行で区切る\n" +
    "- 引用符（\"、『』、「」）は使用しない\n" +
    "- 年齢や不明な情報は推測せず、データに基づいてのみ記載\n" +
    "- 専門用語には初出時に括弧で補足する（例: GE（グリセリン浣腸）、摘便）\n" +
    "- AIらしい前置き（承知しました等）や締めの挨拶は一切不要\n" +
    "- 自然な申し送り調（〜です、〜されています、〜とのことです）で記述\n" +
    "- 同じ語尾の連続を避け、文の長短を混ぜる\n" +
    "- 「記載なし」「不明」「特になし」などの欠落を示す表現は使用しない\n" +
    "- 情報がないセクションはセクション自体を出力しない\n" +
    "- 報告内容に【確認事項】がある場合のみ確認事項セクションを出力し、ない場合は省略する\n" +
    "- 出力の末尾に「記録者: （記録者名）」を必ず記載する";
  
  return callGeminiAPI_(prompt) || "";
}

/**
 * SOAP形式の要約生成（医師用）- Phase 1 改善版
 */
function generateSOAPSummary_(content, patientId, startTime, endTime, reporterName) {
  if (!GEMINI_API_KEY) return "";
  
  var reporterInfo = reporterName ? "記録者: " + reporterName + "\n\n" : "";
  
  var careLevel = getPatientCareLevel_(patientId);
  var needsPsychotherapy = isEligibleForPsychotherapy_(careLevel);
  
  var psychotherapyInstruction = "";
  if (needsPsychotherapy) {
    psychotherapyInstruction = 
      "\n\n診察記録の中で、患者やご家族に対して行った指導や助言がある場合は、" +
      "必ず以下の形式で記載してください：\n\n" +
      "在宅精神療法：\n" +
      "（指導内容を具体的に記載）\n\n" +
      "指導内容が明確でない場合は、在宅生活における一般的な助言として記載してください。";
  }
  
  // 時間情報を事前に整形
  var visitTimeText = "";
  if (startTime && endTime) {
    try {
      var start = new Date(startTime);
      var end = new Date(endTime);
      var duration = Math.round((end - start) / 60000);
      
      var startStr = formatTime_(start);
      var endStr = formatTime_(end);
      
      visitTimeText = "訪問時間: " + startStr + " - " + endStr;
      if (duration > 0) {
        visitTimeText += "\n診察時間: " + duration + "分";
      }
      visitTimeText += "\n\n";
    } catch (e) {
      Logger.log("⚠️ 時間情報の整形エラー: " + e.toString());
    }
  }
  
  var prompt = "あなたは精神科医師です。以下の診察記録から、SOAP形式のカルテ記載を作成してください。\n\n" +
    reporterInfo +
    "診察記録:\n" + content + "\n\n" +
    "出力形式:\n" +
    visitTimeText +
    "S: 患者の主訴、自覚症状、病歴、既往歴、家族歴、生活歴など患者から得られた主観的情報を具体的に記載\n\n" +
    "O: バイタルサイン、身体所見、検査結果など客観的に得られた情報を具体的に記載\n\n" +
    "A: SとOに基づく診断内容、鑑別診断、病態評価を記載\n\n" +
    "P: 治療計画（処方薬、処置、生活指導、検査予定、今後の方針、次回受診日など）を具体的に記載\n" +
    psychotherapyInstruction + "\n\n" +
    "確認事項:\n" +
    "記録者: （記録者名）\n" +
    "診療種別: （訪問/外来）\n" +
    "処方せん変更: （あり/なし）\n" +
    "通院／在宅精神療法: （あり/なし）\n" +
    "※診察記録の【確認事項】に記載されている内容を必ず転記してください。「あり」「なし」のいずれかを明記してください。記載がない場合は両方とも「なし」としてください。\n" +
    "※記録者名は必ず含めてください。\n\n" +
    "厳守事項:\n" +
    "- ですます調は使わず簡潔に記載\n" +
    "- Markdown記法（見出し記号#、太字**）は一切使用禁止\n" +
    "- 箇条書きの記号（・、-、*）は一切使用しない\n" +
    "- 改行のみで項目を区切る\n" +
    "- 引用符（\"、『』、「」）は使用しない\n" +
    "- プレーンテキストのみ\n" +
    "- SOAPの各項目で内容の重複を避ける\n" +
    "- 抽象的な表現（重要、最適、本質等）を避け、具体的な動詞中心の表現にする\n" +
    "- AIらしい前置き（承知しました等）や締めの挨拶は一切不要、本文のみ出力\n" +
    "- 結論から言うと、一概には言えない等のクッション言葉は削除\n" +
    "- 同じ語尾の連続を避け、文の長短を混ぜてリズムを整える\n" +
    "- 推測や不明な情報（年齢等）は記載しない\n" +
    "- 「記載なし」「不明」「情報なし」などの欠落を示す表現は絶対に使用しない\n" +
    "- 情報がない項目は記載せず、ある情報のみを記載する\n" +
    "- 確認事項セクションは必ず出力し、記録者名、診療種別、処方せん変更、通院／在宅精神療法の有無を明記する";
  
  return callGeminiAPI_(prompt) || "";
}

/**
 * 時刻フォーマット用ヘルパー関数
 */
function formatTime_(date) {
  var hours = String(date.getHours());
  if (hours.length === 1) hours = '0' + hours;
  var minutes = String(date.getMinutes());
  if (minutes.length === 1) minutes = '0' + minutes;
  return hours + ':' + minutes;
}

function generateChartText_(content, patientId, patientName, imageUrls) {
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  var text = today + " 訪問看護 " + patientName + "（ID:" + patientId + "）\n" + content;
  
  if (imageUrls) {
    text += "\n\n[添付画像] " + imageUrls.split(',').length + "枚\n" + imageUrls;
  }
  
  return text;
}

function callGeminiAPI_(prompt) {
  if (!GEMINI_API_KEY) {
    Logger.log("❌ エラー: GEMINI_API_KEYが設定されていません");
    return "（要約は後ほど更新されます）";
  }
  
  // 2026年2月時点の最新モデル
  var modelName = "gemini-3-flash-preview"; // 最新のプレビューモデル
  var url = "https://generativelanguage.googleapis.com/v1beta/models/" + modelName + ":generateContent?key=" + GEMINI_API_KEY;
  var payload = { contents: [{ parts: [{ text: prompt }] }] };
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    Logger.log("🤖 Gemini API呼び出し開始");
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    
    Logger.log("📡 レスポンスコード: " + responseCode);
    
    if (responseCode !== 200) {
      Logger.log("❌ APIエラー: " + responseText);
      return "（要約は後ほど更新されます）";
    }
    
    var json = JSON.parse(responseText);
    
    if (!json.candidates || !json.candidates[0] || !json.candidates[0].content) {
      Logger.log("❌ レスポンス形式エラー: " + responseText);
      return "（要約は後ほど更新されます）";
    }
    
    var result = json.candidates[0].content.parts[0].text;
    Logger.log("✅ AI要約生成成功: " + result.substring(0, 50) + "...");
    return result;
    
  } catch (e) {
    Logger.log("❌ 例外エラー: " + e.toString());
    return "（要約は後ほど更新されます）";
  }
}

/**
 * 音声ファイルをGemini APIで文字起こし
 * @param {string} base64Audio - Base64エンコードされた音声データ
 * @param {string} mimeType - 音声ファイルのMIMEタイプ
 * @return {string} - 文字起こしされたテキスト
 */
/**
 * 画像から文字を読み取る（OCR）- 堅牢版
 */
function performOCROnImages(base64Images) {
  try {
    if (!GEMINI_API_KEY) {
      return { success: false, msg: "Gemini APIキーが設定されていません" };
    }
    
    if (!base64Images || base64Images.length === 0) {
      return { success: false, msg: "画像がありません" };
    }
    
    Logger.log("🔍 OCR開始: " + base64Images.length + "枚の画像");
    
    var allOcrText = "";
    var modelName = "gemini-3-flash-preview"; 
    var url = "https://generativelanguage.googleapis.com/v1beta/models/" + modelName + ":generateContent?key=" + GEMINI_API_KEY;

    for (var i = 0; i < base64Images.length; i++) {
      try {
        var base64Image = base64Images[i];
        var base64Clean = base64Image.replace(/^data:image\/\w+;base64,/, '');
        
        var payload = {
          contents: [{
            parts: [
              { text: "この画像は訪問看護の現場で撮影されたものです。処方箋、メモ、あるいはバイタルデータなどが含まれている可能性があります。画像内のすべての文字を、医療記録として利用できるよう正確に読み取ってください。" },
              {
                inline_data: {
                  mime_type: "image/jpeg",
                  data: base64Clean
                }
              }
            ]
          }]
        };

        var options = {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        };

        var response = UrlFetchApp.fetch(url, options);
        var responseCode = response.getResponseCode();
        var responseText = response.getContentText();

        if (responseCode === 200) {
          var json = JSON.parse(responseText);
          if (json.candidates && json.candidates[0] && json.candidates[0].content) {
            var text = json.candidates[0].content.parts[0].text;
            if (text) {
              if (allOcrText) allOcrText += "\n\n--- 次の画像 ---\n\n";
              allOcrText += text;
            }
          }
        } else {
          Logger.log("⚠️ 画像 " + (i + 1) + " のOCRに失敗: " + responseText);
        }
      } catch (innerError) {
        Logger.log("❌ ループ内エラー (画像 " + (i + 1) + "): " + innerError.toString());
      }
    }
    
    return { 
      success: true, 
      ocrText: allOcrText || "" 
    };
  } catch (error) {
    Logger.log("❌ OCR全体エラー: " + error.toString());
    return { success: false, msg: "OCR処理中にエラーが発生しました" };
  }
}

function transcribeAudioWithGemini_(base64Audio, mimeType) {
  if (!GEMINI_API_KEY) {
    Logger.log("❌ エラー: GEMINI_API_KEYが設定されていません");
    return "";
  }
  
  try {
    var modelName = "gemini-3-flash-preview";
    var url = "https://generativelanguage.googleapis.com/v1beta/models/" + modelName + ":generateContent?key=" + GEMINI_API_KEY;
    
    // 音声データとプロンプトを送信
    var payload = {
      contents: [{
        parts: [
          {
            text: "この音声ファイルを文字起こししてください。訪問看護の報告内容です。医療用語や数値は正確に記録してください。"
          },
          {
            inline_data: {
              mime_type: mimeType,
              data: base64Audio
            }
          }
        ]
      }]
    };
    
    var options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    Logger.log("🎤 Gemini API（音声文字起こし）呼び出し開始");
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    
    Logger.log("📡 レスポンスコード: " + responseCode);
    
    if (responseCode !== 200) {
      Logger.log("❌ 音声APIエラー: " + responseText);
      return "";
    }
    
    var json = JSON.parse(responseText);
    
    if (!json.candidates || !json.candidates[0] || !json.candidates[0].content) {
      Logger.log("❌ 音声レスポンス形式エラー: " + responseText);
      return "";
    }
    
    var transcription = json.candidates[0].content.parts[0].text;
    Logger.log("✅ 音声文字起こし成功: " + transcription.substring(0, 100) + "...");
    return transcription;
    
  } catch (e) {
    Logger.log("❌ 音声文字起こしエラー: " + e.toString());
    return "";
  }
}

/**
 * Summaryシートから患者のサマリーを取得（高速）
 * @param {string} patientId - 患者ID
 * @param {string} patientName - 患者名（未使用だが互換性のため残す）
 * @return {object} - {found: boolean, content: string, summary: string, files: array}
 */
function getDocumentSummary_(patientId, patientName) {
  Logger.log("📄 サマリー取得開始: ID=" + patientId);
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var summarySheet = ss.getSheetByName("Summary");
    
    if (!summarySheet) {
      Logger.log("⚠️ Summaryシートが見つかりません");
      return { found: false, content: "", summary: "", files: [] };
    }
    
    // Summaryシートから取得（瞬時に取得）
    var record = getSummaryRecord_(summarySheet, patientId);
    
    if (!record || !record.summary) {
      Logger.log("⚠️ サマリーが見つかりません（ID: " + patientId + "）");
      Logger.log("💡 ヒント: サマリーフォルダにファイルを追加するか、手動で更新してください");
      return { found: false, content: "", summary: "", files: [] };
    }
    
    Logger.log("✅ サマリー取得成功（ファイル数: " + record.fileCount + "）");
    
    var fileNames = record.sourceFiles ? record.sourceFiles.split(", ") : [];
    
    return {
      found: true,
      content: record.summary, // 詳細コンテンツは保存していないため、サマリーを返す
      summary: record.summary,
      files: fileNames
    };
    
  } catch (e) {
    Logger.log("❌ サマリー取得エラー: " + e.toString());
    return { found: false, content: "", summary: "", files: [] };
  }
}

/**
 * PDFファイルからOCRでテキストを抽出
 * @param {File} pdfFile - PDFファイル
 * @return {string} - 抽出されたテキスト
 */
function extractTextFromPDF_(pdfFile) {
  try {
    // 一時的にOCR処理用のGoogleドキュメントを作成
    var resource = {
      title: "temp_ocr_" + new Date().getTime(),
      mimeType: MimeType.GOOGLE_DOCS
    };
    
    var options = {
      ocr: true,
      ocrLanguage: 'ja'
    };
    
    // Drive API v3を使用してPDFをGoogleドキュメントに変換
    var blob = pdfFile.getBlob();
    var tempDoc = Drive.Files.insert(resource, blob, options);
    
    // テキストを抽出
    var doc = DocumentApp.openById(tempDoc.id);
    var text = doc.getBody().getText();
    
    // 一時ドキュメントを削除
    DriveApp.getFileById(tempDoc.id).setTrashed(true);
    
    Logger.log("✅ PDF OCR完了: " + text.length + "文字");
    return text;
    
  } catch (e) {
    Logger.log("❌ PDF OCRエラー: " + e.toString());
    return "[PDFの読み込みに失敗しました: " + e.toString() + "]";
  }
}

/**
 * ============================================
 * サマリー自動生成システム（トリガーベース）
 * ============================================
 */

/**
 * サマリーフォルダを取得（フォルダIDまたはフォルダ名に対応）
 */
function getSummaryFolder_() {
  if (!SUMMARY_FOLDER_NAME) {
    return null;
  }
  
  try {
    // まずフォルダIDとして試す（33文字程度の英数字）
    if (SUMMARY_FOLDER_NAME.length > 20 && SUMMARY_FOLDER_NAME.indexOf(' ') === -1) {
      try {
        var folder = DriveApp.getFolderById(SUMMARY_FOLDER_NAME);
        Logger.log("✅ フォルダをIDで取得: " + folder.getName());
        return folder;
      } catch (e) {
        Logger.log("⚠️ フォルダIDとして取得失敗、フォルダ名として試します");
      }
    }
    
    // フォルダ名として試す
    var folders = DriveApp.getFoldersByName(SUMMARY_FOLDER_NAME);
    if (folders.hasNext()) {
      var folder = folders.next();
      Logger.log("✅ フォルダを名前で取得: " + folder.getName());
      return folder;
    }
    
    return null;
  } catch (e) {
    Logger.log("❌ フォルダ取得エラー: " + e.toString());
    return null;
  }
}

/**
 * Google Driveのサマリーフォルダにファイルが追加されたときのトリガー
 * 注意: このトリガーは手動で設定する必要があります
 * 設定方法: Apps Script > トリガー > トリガーを追加 > イベントのソース: Google Drive
 */
function onFileAddedToSummaryFolder(e) {
  Logger.log("🔔 ファイル追加トリガー発動");
  
  try {
    // トリガーイベントからファイル情報を取得
    var fileId = e && e.fileId ? e.fileId : null;
    
    if (!fileId) {
      Logger.log("⚠️ ファイルIDが取得できません");
      return;
    }
    
    var file = DriveApp.getFileById(fileId);
    var fileName = file.getName();
    Logger.log("📄 追加されたファイル: " + fileName);
    
    // 患者IDを抽出（ファイル名から）
    var patientIds = extractPatientIdsFromFileName_(fileName);
    
    if (patientIds.length === 0) {
      Logger.log("⚠️ ファイル名から患者IDを抽出できませんでした: " + fileName);
      return;
    }
    
    // 各患者のサマリーを更新
    patientIds.forEach(function(patientId) {
      Logger.log("🔄 患者 " + patientId + " のサマリーを更新中...");
      updatePatientSummaryFromDrive_(patientId);
    });
    
  } catch (error) {
    Logger.log("❌ トリガーエラー: " + error.toString());
  }
}

/**
 * ファイル名から患者IDを抽出（厳密マッチング）
 */
function extractPatientIdsFromFileName_(fileName) {
  var ids = [];
  
  // 患者シートから全患者情報を取得
  var pSheet = getPatientSheet_();
  if (!pSheet) return ids;
  
  var data = pSheet.getDataRange().getValues();
  var patients = [];
  
  // 患者情報を配列に格納
  for (var i = 1; i < data.length; i++) {
    var patientId = String(data[i][COL.PATIENT_ID]).trim();
    var patientName = String(data[i][COL.NAME]).trim();
    
    if (patientId) {
      patients.push({
        id: patientId,
        name: patientName,
        idLength: patientId.length
      });
    }
  }
  
  // 患者IDの長い順にソート（部分一致を避けるため）
  patients.sort(function(a, b) {
    return b.idLength - a.idLength;
  });
  
  // ファイル名を正規化（拡張子を除去）
  var fileNameWithoutExt = fileName.replace(/\.(pdf|doc|docx|txt)$/i, '');
  
  // 各患者について厳密にチェック（長い順なので最初にマッチした1人のみ処理）
  for (var i = 0; i < patients.length; i++) {
    var patient = patients[i];
    var matched = false;
    
    // パターン1: 患者IDが完全一致または明確に区切られている
    // 完全一致を優先
    if (fileNameWithoutExt === patient.id) {
      matched = true;
      Logger.log("✅ 完全一致: ファイル「" + fileName + "」→ 患者ID「" + patient.id + "」");
    }
    
    // パターン2: IDが区切り文字で囲まれている（例: "001_山田", "ID-001", "患者001"）
    if (!matched) {
      var idPattern = new RegExp('(^|[^0-9])' + escapeRegExp_(patient.id) + '([^0-9]|$)');
      if (idPattern.test(fileNameWithoutExt)) {
        matched = true;
        Logger.log("✅ ID境界マッチ: ファイル「" + fileName + "」→ 患者ID「" + patient.id + "」");
      }
    }
    
    // パターン3: 患者名が完全一致または区切り文字で囲まれている
    if (!matched && patient.name && patient.name.length >= 2) {
      var namePattern = new RegExp('(^|[_\\-\\s])' + escapeRegExp_(patient.name) + '([_\\-\\s.]|$)');
      if (namePattern.test(fileName)) {
        matched = true;
        Logger.log("✅ 名前マッチ: ファイル「" + fileName + "」→ 患者ID「" + patient.id + "」(" + patient.name + ")");
      }
    }
    
    // 最初にマッチした1人のみ処理してループを抜ける
    if (matched) {
      ids.push(patient.id);
      break; // ★ 重要: 最初の1人のみ
    }
  }
  
  if (ids.length === 0) {
    Logger.log("⚠️ マッチする患者が見つかりません: " + fileName);
  }
  
  return ids;
}

/**
 * 正規表現用のエスケープ処理
 */
function escapeRegExp_(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * 権限の再承認を強制する（テスト用）
 * 実行方法: Apps Script エディタで「実行」→「forceReauthorization」
 */
function forceReauthorization() {
  try {
    // DriveApp、DocumentApp、UrlFetchAppの権限をテスト
    Logger.log("🔐 権限確認を開始します...");
    
    // Drive権限
    DriveApp.getFiles();
    Logger.log("✅ Drive権限: OK");
    
    // Document権限
    var testDoc = DocumentApp.create('権限テスト_' + new Date().getTime());
    var testDocId = testDoc.getId();
    DriveApp.getFileById(testDocId).setTrashed(true);
    Logger.log("✅ Document権限: OK");
    
    // 外部リクエスト権限（Gemini API用）
    var GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (GEMINI_API_KEY) {
      Logger.log("✅ Gemini APIキー: 設定済み");
    } else {
      Logger.log("⚠️ Gemini APIキー: 未設定");
    }
    
    Logger.log("✅ すべての権限確認が完了しました！");
    Logger.log("📝 次のステップ: サマリー更新を実行してください");
    
    return "権限確認成功";
    
  } catch (e) {
    Logger.log("❌ 権限確認エラー: " + e.toString());
    Logger.log("💡 解決方法:");
    Logger.log("1. appsscript.json を確認");
    Logger.log("2. 左側の「サービス」から Drive API (v3) を追加");
    Logger.log("3. この関数を再実行して権限を承認");
    
    throw new Error("権限エラー: " + e.toString());
  }
}

/**
 * 患者のサマリーを既に取得したファイルリストから更新（効率的、逆引き用）
 */
function updatePatientSummaryFromFiles_(patientId, files) {
  Logger.log("📊 患者 " + patientId + " のサマリー更新開始（" + files.length + "ファイル）");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName("Summary");
  
  if (!summarySheet) {
    Logger.log("❌ Summaryシートが見つかりません");
    return;
  }
  
  // 患者名を取得
  var pSheet = getPatientSheet_();
  var patientName = "";
  if (pSheet) {
    var patientData = pSheet.getDataRange().getValues();
    for (var i = 1; i < patientData.length; i++) {
      if (String(patientData[i][COL.PATIENT_ID]).trim() === String(patientId).trim()) {
        patientName = patientData[i][COL.NAME];
        break;
      }
    }
  }
  
  // 既存のサマリーレコードを取得
  var existingRecord = getSummaryRecord_(summarySheet, patientId);
  
  // 手動編集されている場合はスキップ
  if (existingRecord && existingRecord.manuallyEdited) {
    Logger.log("⚠️ 手動編集済みのため、自動更新をスキップします");
    return;
  }
  
  // 処理済みファイルIDリストを取得
  var processedIds = existingRecord ? existingRecord.processedIds.split(",") : [];
  var newFiles = [];
  var allFileNames = [];
  var allFileIds = processedIds.slice(); // 既存のIDを引き継ぐ
  
  // K列のIDリストで重複チェック
  for (var i = 0; i < files.length; i++) {
    var file = files[i];
    var fileId = file.getId();
    var fileName = file.getName();
    allFileNames.push(fileName);
    
    // K列のIDリストに含まれていないファイルのみを「新規」とする
    if (processedIds.indexOf(fileId) === -1) {
      newFiles.push(file);
      allFileIds.push(fileId);
      Logger.log("🆕 新規ファイル: " + fileName + " (ID: " + fileId + ")");
    } else {
      Logger.log("ℹ️ 処理済みファイル（スキップ）: " + fileName);
    }
  }
  
  // 新規ファイルがない場合はスキップ
  if (newFiles.length === 0) {
    Logger.log("ℹ️ 新規ファイルがないため、更新不要です");
    return;
  }
  
  Logger.log("📋 新規ファイル数: " + newFiles.length + " / 合計: " + files.length);
  
  // 新規ファイルのコンテンツを読み込み
  var newContent = [];
  
  for (var i = 0; i < newFiles.length; i++) {
    var file = newFiles[i];
    var fileName = file.getName();
    var mimeType = file.getMimeType();
    
    Logger.log("📄 処理中 [" + (i + 1) + "/" + newFiles.length + "]: " + fileName);
    
    var content = "";
    
    if (mimeType === MimeType.GOOGLE_DOCS) {
      try {
        var doc = DocumentApp.openById(file.getId());
        content = doc.getBody().getText();
      } catch (e) {
        Logger.log("❌ Googleドキュメント読み込みエラー: " + e.toString());
        continue;
      }
    } else if (mimeType === MimeType.PDF) {
      Logger.log("📄 PDFファイルをOCRでテキスト化中...");
      content = extractTextFromPDF_(file);
      
      // PDF読み込みエラーの場合はスキップ
      if (content.indexOf("[PDFの読み込みに失敗しました") !== -1) {
        Logger.log("⚠️ PDFの読み込みに失敗したためスキップします");
        Logger.log("💡 ヒント: Googleドキュメントでテキストをコピペすることをお勧めします");
        continue;
      }
    } else {
      Logger.log("⚠️ スキップ（サポート外形式）: " + mimeType);
      continue;
    }
    
    if (content && content.length > 0 && content.trim() !== "") {
      newContent.push("【" + fileName + "】\n" + content);
      Logger.log("✅ 読み込み完了: " + content.length + "文字");
    } else {
      Logger.log("⚠️ 内容が空のためスキップ: " + fileName);
    }
  }
  
  if (newContent.length === 0) {
    Logger.log("⚠️ 新規コンテンツがありません");
    return;
  }
  
  // サマリーを生成または更新
  var summary = "";
  var combinedNewContent = newContent.join("\n\n─────────────────\n\n");
  
  if (existingRecord && existingRecord.summary) {
    // 既存サマリーがある場合は増分更新
    Logger.log("🔄 既存サマリーに新規情報を統合中...");
    var prompt = "以下は患者の既存サマリーです：\n\n" + existingRecord.summary + 
                 "\n\n新しい情報が追加されました：\n\n" + combinedNewContent +
                 "\n\n既存サマリーに新しい情報を統合して、更新されたサマリー（5〜8行程度）を作成してください。\n\n" +
                 "重要な指示:\n" +
                 "- Markdown記法（**、*、#など）は一切使用しない\n" +
                 "- 箇条書きは「・」のみ使用\n" +
                 "- 年齢や不明な情報を推測しない（データに基づいてのみ記載）\n" +
                 "- 医師への依頼事項（処方・検査など）があれば強調して記載\n" +
                 "- 次回訪問での重点確認事項があれば明記\n" +
                 "- シンプルで読みやすい文章にする";
    summary = callGeminiAPI_(prompt);
  } else {
    // 初回生成
    Logger.log("🆕 新規サマリーを生成中...");
    var prompt = "以下は患者のサマリー文書です。重要なポイントを5〜8行程度に簡潔にまとめてください。\n\n" + 
                 combinedNewContent + "\n\n" +
                 "重要な指示:\n" +
                 "- Markdown記法（**、*、#など）は一切使用しない\n" +
                 "- 箇条書きは「・」のみ使用\n" +
                 "- 年齢や不明な情報を推測しない（データに基づいてのみ記載）\n" +
                 "- 医師への依頼事項（処方・検査など）があれば強調して記載\n" +
                 "- 次回訪問での重点確認事項があれば明記\n" +
                 "- シンプルで読みやすい文章にする";
    summary = callGeminiAPI_(prompt);
  }
  
  // ファイルのリネーム処理：処理済みの印 [S] を付ける
  for (var i = 0; i < newFiles.length; i++) {
    var file = newFiles[i];
    var name = file.getName();
    if (name.indexOf("[S]") === -1) {
      try {
        file.setName(name + " [S]");
        Logger.log("🏷️ ファイルをリネーム: " + name + " → " + name + " [S]");
      } catch (e) {
        Logger.log("⚠️ リネームに失敗: " + name + " (" + e.toString() + ")");
      }
    }
  }
  
  // Summaryシートを更新
  saveSummaryRecord_(summarySheet, {
    patientId: patientId,
    patientName: patientName,
    summary: summary,
    sourceFiles: allFileNames.join(", "),
    fileCount: allFileNames.length,
    previousSummary: existingRecord ? existingRecord.summary : "", // 今までのサマリーをJ列へ移動（バックアップ）
    processedIds: allFileIds.join(","), // 更新されたIDリストをK列へ保存
    manuallyEdited: false
  });
  
  Logger.log("✅ サマリー更新完了");
}

/**
 * 患者のサマリーをDriveファイルから更新（トリガー用）
 */
function updatePatientSummaryFromDrive_(patientId) {
  Logger.log("📊 患者 " + patientId + " のサマリー更新開始（トリガー経由）");
  
  var folder = getSummaryFolder_();
  if (!folder) {
    Logger.log("⚠️ フォルダが見つかりません");
    return;
  }
  
  var allFiles = folder.getFiles();
  var matchedFiles = [];
  
  // 患者名を取得
  var pSheet = getPatientSheet_();
  var patientName = "";
  if (pSheet) {
    var patientData = pSheet.getDataRange().getValues();
    for (var i = 1; i < patientData.length; i++) {
      if (String(patientData[i][COL.PATIENT_ID]).trim() === String(patientId).trim()) {
        patientName = patientData[i][COL.NAME];
        break;
      }
    }
  }
  
  // マッチするファイルを全て取得
  while (allFiles.hasNext()) {
    var file = allFiles.next();
    var fileName = file.getName();
    
    if (fileName.indexOf(patientId) !== -1 || 
        (patientName && fileName.indexOf(patientName) !== -1)) {
      matchedFiles.push(file);
    }
  }
  
  if (matchedFiles.length === 0) {
    Logger.log("⚠️ マッチするファイルが見つかりません");
    return;
  }
  
  // ファイルリストを使って更新（既存の処理ロジックを再利用）
  updatePatientSummaryFromFiles_(patientId, matchedFiles);
}

/**
 * Summaryシートから患者のレコードを取得
 */
/**
 * Summaryシートから患者のレコードを取得（更新版）
 */
function getSummaryRecord_(summarySheet, patientId) {
  var data = summarySheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(patientId).trim()) {
      return {
        row: i + 1,
        patientId: data[i][0],
        patientName: data[i][1],
        lastUpdated: data[i][2],
        summary: data[i][3],
        sourceFiles: data[i][4],
        manuallyEdited: data[i][5] === true || data[i][5] === "TRUE",
        manualEditDate: data[i][6],
        lastReportId: data[i][7],
        fileCount: data[i][8],
        previousSummary: data[i][9] || "",
        processedIds: String(data[i][10] || ""),
        recentUpdates: data[i][11] || "",
        updateCount: data[i][12] || 0
      };
    }
  }
  
  return null;
}

/**
 * Summaryシートにレコードを保存
 */
/**
 * Summaryシートにレコードを保存（更新版）
 */
function saveSummaryRecord_(summarySheet, record) {
  var existingRecord = getSummaryRecord_(summarySheet, record.patientId);
  var now = new Date();
  
  // 保存用データ配列（13項目）
  var rowData = [
    record.patientId,
    record.patientName,
    now,
    record.summary,
    record.sourceFiles,
    record.manuallyEdited || false,
    record.manualEditDate || "",
    record.lastReportId || "",
    record.fileCount || 0,
    record.previousSummary || "",
    record.processedIds || "",
    record.recentUpdates || "",
    record.updateCount || 0
  ];
  
  if (existingRecord) {
    summarySheet.getRange(existingRecord.row, 1, 1, rowData.length).setValues([rowData]);
    Logger.log("💾 サマリーを更新しました（行" + existingRecord.row + "）");
  } else {
    summarySheet.appendRow(rowData);
    Logger.log("💾 サマリーを新規作成しました");
  }
}

/**
 * 要約サマリーに医師への依頼が含まれる場合にメールを送信する
 */
function sendDoctorAlertIfRequested_(summary, patientId, patientName, reporter) {
  var header = "【★医師への依頼】";
  if (!summary.includes(header)) return;
  
  // 依頼セクションの内容を抽出
  var parts = summary.split(header);
  var requestText = parts[1].split("【")[0].trim(); // 次の【セクションまでを切り出し
  
  // 「なし」と書かれている場合は送信しない
  if (requestText === "" || requestText === "なし" || requestText.indexOf("なし") === 0) {
    return;
  }
  
  var recipient = "ishihara.brain@gmail.com";
  var subject = "【要確認】訪問看護より医師への依頼（" + patientName + "様）";
  
  var body = "以下の患者様の報告にて、医師への依頼事項が発生しました。\n\n" +
             "■ 患者ID: " + patientId + "\n" +
             "■ 患者名: " + patientName + " 様\n" +
             "■ 報告者: " + reporter + "\n\n" +
             "--------------------------------------------\n" +
             "【依頼内容】\n" +
             requestText + "\n" +
             "--------------------------------------------\n\n" +
             "詳細は管理システムまたはスプレッドシートをご確認ください。";
             
  try {
    MailApp.sendEmail(recipient, subject, body);
    Logger.log("📧 医師へのアラートメールを送信しました: " + patientName);
  } catch (e) {
    Logger.log("❌ メール送信エラー: " + e.toString());
  }
}

/**
 * 緊急・訪問前連絡を直接メール送信する
 */
function sendUrgentDoctorAlert(patientId, patientName, reporter, content) {
  var recipient = "ishihara.brain@gmail.com";
  var subject = "【至急・訪問前】連携システムより直接連絡（" + patientName + "様）";
  var body = "医師へ至急伝えたい内容が送信されました。\n\n" +
             "■ 患者名: " + patientName + " 様\n" +
             "■ 連絡者: " + reporter + "\n" +
             "--------------------------------------------\n" +
             "【連絡内容】\n" + content + "\n" +
             "--------------------------------------------";
  
  try {
    MailApp.sendEmail(recipient, subject, body);
    Logger.log("🚀 緊急メールを送信しました: " + patientName);
    return { success: true };
  } catch (e) {
    Logger.log("❌ 緊急メール送信エラー: " + e.toString());
    return { success: false, msg: e.toString() };
  }
}

/**
 * 医師の診察記録から次回タスクを抽出してリマインドメールを送信
 */
function sendDoctorReminderEmail(patientId, patientName, reporter, content) {
  if (!GEMINI_API_KEY) {
    Logger.log("⚠️ APIキーが未設定のため、リマインド抽出をスキップ");
    return { success: false, msg: "APIキーが未設定です" };
  }
  
  var prompt = "以下の診察記録から、医師が次回行うべきタスクを抽出してください。\n\n" +
    "診察記録:\n" + content + "\n\n" +
    "抽出対象:\n" +
    "紹介状の作成\n" +
    "処方箋の変更\n" +
    "検査の指示\n" +
    "書類の記載\n" +
    "他院への連絡\n" +
    "その他の要対応事項\n\n" +
    "出力形式:\n" +
    "タスクがある場合は改行のみで列挙してください。\n" +
    "タスクがない場合は「なし」と出力してください。\n\n" +
    "厳守事項:\n" +
    "- 箇条書きの記号（・、-、*）は使用しない\n" +
    "- 改行のみで区切る\n" +
    "- 前置きや締めの挨拶は不要\n" +
    "- 具体的で簡潔に記載";
  
  var tasks = callGeminiAPI_(prompt);
  
  // タスクがない場合は送信しない
  if (!tasks || tasks.trim() === "" || tasks.trim() === "なし") {
    Logger.log("ℹ️ 次回タスクなし、リマインドメール送信をスキップ");
    return { success: true, skipped: true };
  }
  
  var recipient = "ishihara.brain@gmail.com";
  var subject = "【医師のリマインド】次回対応タスク（" + patientName + "様）";
  
  var body = "診察記録から次回行うべきタスクを抽出しました。\n\n" +
             "■ 患者名: " + patientName + " 様\n" +
             "■ 患者ID: " + patientId + "\n" +
             "■ 記録者: " + reporter + "\n" +
             "■ 記録日時: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm") + "\n\n" +
             "--------------------------------------------\n" +
             "【次回対応タスク】\n" +
             tasks + "\n" +
             "--------------------------------------------\n\n" +
             "詳細は管理システムまたはスプレッドシートをご確認ください。";
             
  try {
    MailApp.sendEmail(recipient, subject, body);
    Logger.log("📧 医師リマインドメールを送信しました: " + patientName);
    return { success: true };
  } catch (e) {
    Logger.log("❌ リマインドメール送信エラー: " + e.toString());
    return { success: false, msg: e.toString() };
  }
}

/**
 * 患者の要介護度を取得
 */
function getPatientCareLevel_(patientId) {
  var pSheet = getPatientSheet_();
  if (!pSheet) return "";
  
  var data = pSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL.PATIENT_ID]).trim() === String(patientId).trim()) {
      return data[i][COL.CARE_LEVEL] || "";
    }
  }
  return "";
}

/**
 * 在宅精神療法の対象かチェック（要介護2以上）
 */
function isEligibleForPsychotherapy_(careLevel) {
  var level = String(careLevel).trim();
  return level === "要介護2" || level === "要介護3" || 
         level === "要介護4" || level === "要介護5";
}

/**
 * AI要約を手動修正して保存
 */
function updateManualSummary(reportId, editedSummary, editorName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Reports");
    if (!sheet) return { success: false, msg: "Reportsシートが見つかりません" };
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === reportId) {
        var now = new Date();
        var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
        var editor = editorName || "不明";
        
        sheet.getRange(i + 1, 12).setValue(editedSummary);  // L列: AI要約（修正後）
        sheet.getRange(i + 1, 18).setValue(true);             // R列: 手動修正フラグ
        sheet.getRange(i + 1, 19).setValue(now);              // S列: 修正日時
        sheet.getRange(i + 1, 20).setValue(editor);           // T列: 修正者名
        
        // U列: 修正歴ログ（追記形式）
        var existingLog = String(data[i][20] || "");
        var newLogEntry = dateStr + " " + editor + ": サマリーを修正";
        var updatedLog = existingLog ? existingLog + "\n" + newLogEntry : newLogEntry;
        sheet.getRange(i + 1, 21).setValue(updatedLog);       // U列: 修正歴ログ
        
        Logger.log("✅ 手動修正を保存しました: 行" + (i + 1) + " 修正者: " + editor);
        return { success: true };
      }
    }
    
    return { success: false, msg: "該当する報告が見つかりません" };
  } catch (error) {
    Logger.log("❌ 手動修正保存エラー: " + error.toString());
    return { success: false, msg: error.toString() };
  }
}

/**
 * 手動修正を踏まえてサマリーを再生成
 */
function regenerateSummaryWithContext(reportId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Reports");
    if (!sheet) return { success: false, msg: "Reportsシートが見つかりません" };
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === reportId) {
        var content = data[i][7];
        var currentSummary = data[i][11];
        var reporterRole = data[i][14] || "nurse";
        var patientId = data[i][2];
        var startTime = data[i][15] || "";
        var endTime = data[i][16] || "";
        
        var newSummary = "";
        
        if (reporterRole === 'doctor') {
          newSummary = regenerateSOAPWithFeedback_(content, currentSummary, patientId, startTime, endTime);
        } else {
          newSummary = regenerateNurseSummaryWithFeedback_(content, currentSummary);
        }
        
        sheet.getRange(i + 1, 12).setValue(newSummary);
        sheet.getRange(i + 1, 18).setValue(false);
        
        Logger.log("✅ サマリーを再生成しました: 行" + (i + 1));
        return { success: true, summary: newSummary };
      }
    }
    
    return { success: false, msg: "該当する報告が見つかりません" };
  } catch (error) {
    Logger.log("❌ 再生成エラー: " + error.toString());
    return { success: false, msg: error.toString() };
  }
}

/**
 * SOAP形式の再生成（修正内容を考慮）
 */
function regenerateSOAPWithFeedback_(content, previousSummary, patientId, startTime, endTime) {
  if (!GEMINI_API_KEY) return "";
  
  var careLevel = getPatientCareLevel_(patientId);
  var needsPsychotherapy = isEligibleForPsychotherapy_(careLevel);
  
  var psychotherapyInstruction = "";
  if (needsPsychotherapy) {
    psychotherapyInstruction = 
      "\n\n診察記録の中で、患者やご家族に対して行った指導や助言がある場合は、" +
      "必ず以下の形式で記載してください：\n\n" +
      "在宅精神療法：\n" +
      "（指導内容を具体的に記載）\n\n" +
      "指導内容が明確でない場合は、在宅生活における一般的な助言として記載してください。";
  }
  
  var visitTimeText = "";
  if (startTime && endTime) {
    try {
      var start = new Date(startTime);
      var end = new Date(endTime);
      var duration = Math.round((end - start) / 60000);
      
      var startStr = formatTime_(start);
      var endStr = formatTime_(end);
      
      visitTimeText = "訪問時間: " + startStr + " - " + endStr;
      if (duration > 0) {
        visitTimeText += "\n診察時間: " + duration + "分";
      }
      visitTimeText += "\n\n";
    } catch (e) {
      Logger.log("⚠️ 時間情報の整形エラー: " + e.toString());
    }
  }
  
  var prompt = "あなたは日本語のプロ編集者かつ経験豊富な医師です。以下の診察記録から、SOAP形式のカルテ記載を作成してください。\n\n" +
    "前回生成したサマリー（参考）:\n" + previousSummary + "\n\n" +
    "このサマリーが手動修正されています。修正された内容の意図を汲み取り、同じ方向性で新しいサマリーを作成してください。\n\n" +
    "診察記録:\n" + content + "\n\n" +
    "出力形式:\n" +
    visitTimeText +
    "S: 患者の主訴、自覚症状、病歴、既往歴、家族歴、生活歴など患者から得られた主観的情報を具体的に記載\n\n" +
    "O: バイタルサイン、身体所見、検査結果など客観的に得られた情報を具体的に記載\n\n" +
    "A: SとOに基づく診断内容、鑑別診断、病態評価を記載\n\n" +
    "P: 治療計画（処方薬、処置、生活指導、検査予定、今後の方針、次回受診日など）を具体的に記載\n" +
    psychotherapyInstruction + "\n\n" +
    "厳守事項:\n" +
    "- ですます調は使わず簡潔に記載\n" +
    "- Markdown記法（見出し記号#、太字**）は一切使用禁止\n" +
    "- 箇条書きの記号（・、-、*）は一切使用しない\n" +
    "- 改行のみで項目を区切る\n" +
    "- 引用符（\"、『』、「」）は使用しない\n" +
    "- プレーンテキストのみ\n" +
    "- SOAPの各項目で内容の重複を避ける\n" +
    "- 抽象的な表現（重要、最適、本質等）を避け、具体的な動詞中心の表現にする\n" +
    "- AIらしい前置き（承知しました等）や締めの挨拶は一切不要、本文のみ出力\n" +
    "- 結論から言うと、一概には言えない等のクッション言葉は削除\n" +
    "- 同じ語尾の連続を避け、文の長短を混ぜてリズムを整える\n" +
    "- 推測や不明な情報（年齢等）は記載しない\n" +
    "- 「記載なし」「不明」「情報なし」などの欠落を示す表現は絶対に使用しない\n" +
    "- 情報がない項目は記載せず、ある情報のみを記載する";
  
  return callGeminiAPI_(prompt) || "";
}

/**
 * 看護記録の再生成（修正内容を考慮）
 */
function regenerateNurseSummaryWithFeedback_(content, previousSummary) {
  if (!GEMINI_API_KEY) return "";
  
  var prompt = "あなたは経験豊富な訪問看護師です。以下の報告を読み、他のスタッフが見て分かりやすい申し送りを作成してください。\n\n" +
    "前回生成したサマリー（参考）:\n" + previousSummary + "\n\n" +
    "このサマリーが手動修正されています。修正された内容の意図を汲み取り、同じ方向性で新しいサマリーを作成してください。\n\n" +
    "報告内容:\n" + content + "\n\n" +
    "出力形式（項目名は【 】で囲み、各項目は改行で区切る）:\n\n" +
    "【バイタルサイン】\n" +
    "バイタルデータがある場合、数値を一字一句変えずにそのまま記載する。\n\n" +
    "【本日の状態】\n" +
    "患者の主訴や様子を2〜3文で記述する。\n\n" +
    "【栄養・水分】\n" +
    "飲水量、食事内容など栄養関連の情報がある場合に記載する。\n\n" +
    "【排泄】\n" +
    "排便・排尿に関する情報がある場合に記載する。\n\n" +
    "【ケア実施内容】\n" +
    "実施したケアがある場合に記載する。\n\n" +
    "【特記事項】\n" +
    "注意事項があれば記載する。なければ省略する。\n\n" +
    "【医師への依頼】\n" +
    "処方変更・検査依頼などがあれば明記する。なければ省略する。\n\n" +
    "【次回訪問】\n" +
    "次回訪問日時と確認すべき重点項目を記載する。\n\n" +
    "最重要ルール（数値の正確な転記）:\n" +
    "報告内容に含まれる数値データ（体温、血圧、脈拍、SPO2、呼吸数、飲水量、日付、時刻など）は絶対に変更・丸め・省略しない。報告にある数値をそのまま転記すること。\n\n" +
    "厳守事項:\n" +
    "- Markdown記法（見出し記号#、太字**）は一切使用禁止\n" +
    "- 箇条書きの記号（・、-、*、ハイフン）は使用しない\n" +
    "- 項目名は必ず【 】で囲む\n" +
    "- 引用符（\"、『』、「」）は使用しない\n" +
    "- 専門用語には初出時に括弧で補足する\n" +
    "- AIらしい前置き（承知しました等）や締めの挨拶は一切不要\n" +
    "- 自然な申し送り調（〜です、〜されています、〜とのことです）で記述\n" +
    "- 同じ語尾の連続を避け、文の長短を混ぜる\n" +
    "- 「記載なし」「不明」「特になし」などの欠落を示す表現は使用しない\n" +
    "- 情報がないセクションはセクション自体を出力しない";
  
  return callGeminiAPI_(prompt) || "";
}

/**
 * 患者のサマリーを取得（表示用）- 柔軟版
 */
function getSummaryForPatient(patientId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var summarySheet = ss.getSheetByName("Summary");
    
    if (!summarySheet) {
      return { success: false, msg: "Summaryシートが見つかりません" };
    }
    
    var record = getSummaryRecord_(summarySheet, patientId);
    
    if (!record) {
      return { 
        success: true, 
        summary: "",
        recentOnly: false,
        isEmpty: true
      };
    }
    
    var hasBackground = record.summary && record.summary.trim() !== "";
    var hasRecentUpdates = record.recentUpdates && record.recentUpdates.trim() !== "";
    
    // ケース1: 両方ある → 統合して返す
    if (hasBackground && hasRecentUpdates) {
      var integrated = integrateBackgroundAndRecent_(record.summary, record.recentUpdates);
      return {
        success: true,
        summary: integrated,
        recentOnly: false,
        lastUpdated: record.lastUpdated,
        updateCount: record.updateCount
      };
    }
    
    // ケース2: 背景情報のみ → そのまま返す
    if (hasBackground && !hasRecentUpdates) {
      return {
        success: true,
        summary: record.summary,
        recentOnly: false,
        lastUpdated: record.lastUpdated,
        updateCount: record.updateCount
      };
    }
    
    // ケース3: 最近の経過のみ → Recent_Updatesから整形して返す
    if (!hasBackground && hasRecentUpdates) {
      var recentSummary = formatRecentUpdatesOnly_(record.recentUpdates);
      return {
        success: true,
        summary: recentSummary,
        recentOnly: true,
        lastUpdated: record.lastUpdated,
        updateCount: record.updateCount
      };
    }
    
    // ケース4: 両方ない
    return { 
      success: true, 
      summary: "",
      recentOnly: false,
      isEmpty: true
    };
    
  } catch (e) {
    Logger.log("サマリー取得エラー: " + e.toString());
    return { success: false, msg: e.toString() };
  }
}

/**
 * 背景情報と最近の経過を統合
 */
function integrateBackgroundAndRecent_(background, recentUpdates) {
  var recentLines = recentUpdates.split('\n');
  var recentSummary = [];
  
  for (var i = 0; i < recentLines.length; i++) {
    var line = recentLines[i].trim();
    if (line && !line.startsWith('---')) {
      recentSummary.push(line);
    }
  }
  
  var integrated = background + "\n\n" +
    "【最近の経過】\n" +
    recentSummary.slice(-5).join('\n');
  
  return integrated;
}

/**
 * Recent_Updatesのみを整形して表示用にする
 */
function formatRecentUpdatesOnly_(recentUpdates) {
  var lines = recentUpdates.split('\n');
  var formattedLines = [];
  
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (line && !line.startsWith('---')) {
      formattedLines.push(line);
    }
  }
  
  var recentItems = formattedLines.slice(-10);
  
  return "【最近の訪問記録より】\n" + recentItems.join('\n');
}

/**
 * 更新ログを構築（Recent_Updates用）- 改善版
 */
function buildUpdateLog_(newReports) {
  var log = "";
  
  newReports.forEach(function(report) {
    var dateStr = Utilities.formatDate(report.date, Session.getScriptTimeZone(), "MM/dd");
    
    var cleanReporter = String(report.reporter || "").replace(/^患者本人_/, '').replace(/\[医師\]|\[看護\]/g, '');
    
    var summary = (report.aiSummary || "").replace(/\n/g, ' ').substring(0, 80);
    if ((report.aiSummary || "").length > 80) summary += "...";
    
    log += dateStr + " (" + cleanReporter + "): " + summary + "\n";
  });
  
  return log;
}

/**
 * 増分更新のプロンプトを構築
 */
function buildUpdatePrompt_(record, newReports) {
  var prompt = "あなたは訪問診療のベテラン看護師です。";
  
  var hasBackground = record.summary && record.summary.trim() !== "";
  
  if (hasBackground) {
    prompt += "患者の既存サマリーに、新しい訪問記録の重要な情報を統合してください。\n\n";
    prompt += "【現在のサマリー】\n" + record.summary + "\n\n";
  } else {
    prompt += "以下の訪問記録から、患者の最近の状態をまとめたサマリーを作成してください。\n\n";
    prompt += "【注意】背景情報はまだありません。今ある訪問記録のみから、最近の状態を簡潔にまとめてください。\n\n";
  }
  
  prompt += "【新しい訪問記録（" + newReports.length + "件）】\n";
  newReports.forEach(function(report, index) {
    var dateStr = Utilities.formatDate(report.date, Session.getScriptTimeZone(), "MM/dd HH:mm");
    prompt += "---\n";
    prompt += dateStr;
    if (report.scenario) prompt += " | " + report.scenario;
    if (report.manualEdit) prompt += " 【手動修正済み】";
    prompt += "\n";
    
    var cleanReporter = String(report.reporter || "").replace(/^患者本人_/, '');
    if (cleanReporter) {
      prompt += "記録者: " + cleanReporter + "\n";
    }
    
    prompt += report.aiSummary + "\n";
  });
  
  if (hasBackground) {
    prompt += "\n【統合方法】\n" +
      "1. 既存サマリーの基本情報・背景は維持する\n" +
      "2. 新しい訪問記録から、以下の重要な情報を抽出して追記する:\n" +
      "   - 症状の変化（改善・悪化）\n" +
      "   - 新たな問題点\n" +
      "   - 治療方針の変更\n" +
      "   - 家族状況の変化\n" +
      "   - 特筆すべき出来事\n" +
      "3. 既存の情報と矛盾する場合は、新しい情報を優先する\n" +
      "4. 全体として、時系列で自然な流れのあるサマリーにする\n" +
      "5. 重複する内容は統合し、簡潔にまとめる\n";
  } else {
    prompt += "\n【作成方法】\n" +
      "1. 上記の訪問記録から、患者の最近の状態を把握する\n" +
      "2. 以下の観点で情報を整理する:\n" +
      "   - 全体的な状態（安定・変動・悪化など）\n" +
      "   - 主な症状や訴え\n" +
      "   - 現在の治療内容\n" +
      "   - 注意すべき点\n" +
      "3. 時系列で変化がわかるように記述する\n" +
      "4. 3〜5文程度で簡潔にまとめる\n";
  }
  
  prompt += "\n【厳守事項】\n" +
    "- 箇条書きの記号（・、-、*）は使用しない\n" +
    "- 引用符（\"、『』、「」）は使用しない\n" +
    "- 「記載なし」「不明」などのネガティブ表現は使用しない\n" +
    "- AIらしい前置きや締めの挨拶は不要\n" +
    "- 自然な申し送り調（〜です、〜されています）で記述\n" +
    "- 段落分けは適切に行い、読みやすくする\n";
  
  if (hasBackground) {
    prompt += "- 全体で5〜8文程度に収める\n";
  } else {
    prompt += "- 全体で3〜5文程度に収める\n";
  }
  
  return prompt;
}

/**
 * 報告の修正歴を取得
 */
function getEditHistory(reportId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Reports");
    if (!sheet) return { success: false, msg: "Reportsシートが見つかりません" };
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === reportId) {
        var editLog = String(data[i][20] || "");
        var manualEdit = data[i][17] || false;
        var lastEditDate = data[i][18];
        var lastEditor = data[i][19] || "";
        
        var lastEditDateStr = "";
        if (lastEditDate instanceof Date) {
          lastEditDateStr = Utilities.formatDate(lastEditDate, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
        }
        
        return {
          success: true,
          manualEdit: manualEdit,
          lastEditDate: lastEditDateStr,
          lastEditor: lastEditor,
          editLog: editLog
        };
      }
    }
    
    return { success: false, msg: "該当する報告が見つかりません" };
  } catch (error) {
    Logger.log("❌ 修正歴取得エラー: " + error.toString());
    return { success: false, msg: error.toString() };
  }
}

/**
 * サマリーに追記する（修正歴を保持したまま追記）
 */
function appendToSummary(reportId, appendText, editorName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Reports");
    if (!sheet) return { success: false, msg: "Reportsシートが見つかりません" };
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === reportId) {
        var now = new Date();
        var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
        var editor = editorName || "不明";
        
        // 既存のサマリーに追記
        var currentSummary = String(data[i][11] || "");
        var appendedSummary = currentSummary + "\n\n【追記 " + dateStr + " " + editor + "】\n" + appendText;
        
        sheet.getRange(i + 1, 12).setValue(appendedSummary); // L列: AI要約（追記後）
        sheet.getRange(i + 1, 18).setValue(true);             // R列: 手動修正フラグ
        sheet.getRange(i + 1, 19).setValue(now);              // S列: 修正日時
        sheet.getRange(i + 1, 20).setValue(editor);           // T列: 修正者名
        
        // U列: 修正歴ログ（追記）
        var existingLog = String(data[i][20] || "");
        var newLogEntry = dateStr + " " + editor + ": 追記";
        var updatedLog = existingLog ? existingLog + "\n" + newLogEntry : newLogEntry;
        sheet.getRange(i + 1, 21).setValue(updatedLog);       // U列: 修正歴ログ
        
        Logger.log("✅ 追記を保存しました: 行" + (i + 1) + " 追記者: " + editor);
        return { success: true, summary: appendedSummary };
      }
    }
    
    return { success: false, msg: "該当する報告が見つかりません" };
  } catch (error) {
    Logger.log("❌ 追記保存エラー: " + error.toString());
    return { success: false, msg: error.toString() };
  }
}
