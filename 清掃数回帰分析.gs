function regression_analysis() {
  const ss = SpreadsheetApp.openById("1-Gj8-asRGiRA_-S8JHqbZumP_Cu-WlU6bdtiYr5p6Io");
  const sheet = ss.getSheetByName("直近3か月増加率");

  // 最終行を取得
  const lastRow = sheet.getLastRow();

  // A列（1列目）とD列（4列目）を取得
  const rangeA = sheet.getRange(2, 1, lastRow - 1, 1); // A列
  const rangeD = sheet.getRange(2, 4, lastRow - 1, 1); // D列

  const dataA = rangeA.getValues(); // A列データ
  const dataD = rangeD.getValues(); // D列データ

  // A列を文字列に変換し、A列とD列を結合
  const raw_data = dataA.map((valueA, index) => {
    // A列のデータを文字列として変換
    const stringData = valueA[0] instanceof Date 
      ? Utilities.formatDate(valueA[0], Session.getScriptTimeZone(), "yyyyMMdd") // Date型ならフォーマット
      : String(valueA[0]); // それ以外はそのまま文字列化

    // A列とD列を結合
    return [stringData, dataD[index][0]];
  }).filter(row => row[0] !== "" && row[1] !== "" && row[0] !== null && row[1] !== null); // 空白を除外

  // デバッグログ
  Logger.log(raw_data);

  const url = "https://asia-northeast1-m2m-core.cloudfunctions.net/cleaningRegression";
  const payload = {
    raw_data: raw_data,
    future_days: 90
  };

  // HTTP リクエストのオプションを設定
  const options = {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  try {
    // HTTP リクエストを送信
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    Logger.log("Predicted Rates: " + result.predicted_rates);
    Logger.log("Future Dates: " + result.future_dates);
    Logger.log("Slope: " + result.slope);
    Logger.log("Intercept: " + result.intercept);
    Logger.log("R² Score: " + result.r_squared);

    // 結果を "3か月先予測" シートに出力
    const predictionSheet = ss.getSheetByName("3か月先予測");
    if (!predictionSheet) {
      throw new Error("Sheet named '3か月先予測' not found!");
    }

    // Future Dates を yyyy/mm/dd に変換
    const futureDates = result.future_dates.map(date => {
      const [year, month, day] = date.split("-");
      return [`${year}/${month}/${day}`]; // yyyy/mm/dd 形式に整形
    });

    // Predicted Rates を縦方向に整形
    const predictedRates = result.predicted_rates.map(rate => [rate]);

    // Future Dates と Predicted Rates を1列目と2列目に書き込む
    const combinedData = futureDates.map((date, index) => [date[0], predictedRates[index][0]]);
    predictionSheet.getRange(1, 1, combinedData.length, 2).setValues(combinedData);

    return result;
  } catch (error) {
    Logger.log("Error: " + error.toString());
    throw new Error("Failed to call Cloud Function: " + error.toString());
  }
}
