function scrapeSuumoData() {
  // スクレイピング対象のURL
  var baseUrl =
    "https://suumo.jp/jj/bukken/ichiran/JJ012FC001/?ar=030&bs=011&cn=9999999&cnb=0&ekTjCd=&ekTjNm=&kb=1&kt=9999999&mb=0&mt=9999999&sc=13224&ta=13&tj=0&pc=100&po=8&pj=2";

  // 新しいシート名を作成（タイムスタンプを使用）
  var sheetName =
    "多摩市_" + Utilities.formatDate(new Date(), "GMT+9", "yyyyMMdd");

  // すでに同じ名前のシートが存在するか確認
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var existingSheet = spreadsheet.getSheetByName(sheetName);

  if (existingSheet) {
    // シートが存在する場合はそのシートを取得
    var newSheet = existingSheet;
    // シートの内容をクリア（上書き更新）
    newSheet.clear();
  } else {
    // シートが存在しない場合は新しいシートを作成
    var newSheet = spreadsheet.insertSheet(sheetName);
  }

  // ヘッダーを書き込む
  newSheet
    .getRange(1, 1, 1, 12)
    .setValues([
      [
        null,
        null,
        "物件名",
        "価格(万円)",
        "専有面積(m2)",
        "築年数",
        "間取り",
        "沿線",
        "駅",
        "徒歩(分)",
        "所在地",
        "URL",
      ],
    ]);

  // ページ数を取得
  var pageCount = getPageCount(baseUrl);

  // 正規表現
  var regex_trainLine = /([^「]+)「/; //沿線
  var regex_station = /「([^」]+)」/; //駅
  var regex_walkingTime = /歩(\d+)分/; //徒歩(分)

  // 各ページにアクセスして物件情報を取得
  for (var page = 1; page <= pageCount; page++) {
    console.log("========== " + page + "/" + pageCount + "ページ目 ==========");
    // ページごとのURLを構築
    var url = page === 1 ? baseUrl : baseUrl + "&page=" + page;

    // URLからHTMLを取得
    var html = UrlFetchApp.fetch(url).getContentText();

    // Cheerioを使用してHTMLを解析
    var $ = Cheerio.load(html);

    // 物件情報を取得して書き込む
    $(".property_unit").each(function (index, element) {
      // console.log(index + 1 + "件目");

      try {
        // 沿線・駅の情報を取得する
        var stationInfo = $(element)
          .find(".dottable-line dl:contains('沿線・駅') dd")
          .contents()
          .first()
          .text()
          .trim();

        // 取得した値を各項目に設定する
        var propertyInfo = {
          name: $(element)
            .find(".dottable-line dl:contains('物件名') dd")
            .contents()
            .first()
            .text()
            .trim(),
          price: convertPriceString(
            $(element)
              .find(".dottable-line dl:contains('販売価格') span")
              .contents()
              .first()
              .text()
              .trim()
          ),
          location: $(element)
            .find(".dottable-line dl:contains('所在地') dd")
            .contents()
            .first()
            .text()
            .trim(),
          trainLine: stationInfo.match(regex_trainLine)[1],
          station: stationInfo.match(regex_station)[1],
          walkingTime: stationInfo.match(regex_walkingTime)[1],
          space: $(element)
            .find(".dottable-line dl:contains('専有面積') dd")
            .contents()
            .first()
            .text()
            .trim()
            .replace("m", ""),
          layout: $(element)
            .find('.dottable-line dl:contains("間取り") dd')
            .contents()
            .first()
            .text()
            .trim(),
          age: $(element)
            .find(".dottable-line dl:contains('築年月') dd")
            .contents()
            .first()
            .text()
            .trim(),
          url:
            "https://suumo.jp" +
            $(element).find(".property_unit-title a").attr("href"),
        };

        // データを新しいシートに書き込む
        newSheet.appendRow([
          null,
          null,
          propertyInfo.name,
          propertyInfo.price,
          propertyInfo.space,
          propertyInfo.age,
          propertyInfo.layout,
          propertyInfo.trainLine,
          propertyInfo.station,
          propertyInfo.walkingTime,
          propertyInfo.location,
          propertyInfo.url,
        ]);
      } catch (error) {
        Logger.log(index + 1 + "件目", "Error processing property:", error); // エラーログ
      }
    });
  }
}

// ページ数を取得する関数
function getPageCount(baseUrl) {
  var html = UrlFetchApp.fetch(baseUrl).getContentText();
  var $ = Cheerio.load(html);
  var lastPageElement = $(".pagination_set-nav")
    .first()
    .find("li:last-child a");

  if (lastPageElement.length > 0) {
    return parseInt(lastPageElement.text(), 10);
  }
  return 1;
}

// 販売価格を変換する関数
function convertPriceString(priceString) {
  // 末尾が"億"で終わる場合は、"億"とその前にある数字を取り出して数値に変換
  // var matchBillion = priceString.match(/(\d+)億(\d+)万円$/);
  var matchBillion = priceString.match(/^(\d+)億/);
  var billionPart = matchBillion ? parseInt(matchBillion[1], 10) * 10000 : 0;

  // 末尾が"万円"で終わる場合は、万の部分を取り出して数値に変換
  var matchTenThousand = priceString.match(/(\d+)万円$/);
  var tenThousandPart = matchTenThousand
    ? parseInt(matchTenThousand[1], 10)
    : 0;

  // 億と万を合算して返す
  var result = (billionPart + tenThousandPart).toString();

  return result;
}
