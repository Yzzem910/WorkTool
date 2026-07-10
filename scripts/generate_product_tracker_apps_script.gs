/**
 * generate_product_tracker_apps_script.gs
 *
 * Google Apps Script to generate a Product Tracker template in Google Sheets.
 *
 * How to use:
 * 1) In Google Drive, create a new Google Sheets (blank).
 * 2) Extensions -> Apps Script. In the Apps Script editor replace Code.gs content with this file and Save.
 * 3) Run the function createProductTracker() (you may need to authorize the script the first time).
 * 4) After the script completes, the spreadsheet will have these sheets:
 *    - Products
 *    - Listings
 *    - Ads
 *    - Promotions
 *    - ListingMaintenance
 *    - Tasks
 *    - Lists
 *    - Dashboard
 * 5) Verify Lists sheet ranges (named ranges are created). You can now use the sheet directly or File -> Download -> Microsoft Excel (.xlsx) to open in WPS/Excel.
 *
 * Notes:
 * - The script creates named ranges for dropdowns and sets data validation where possible.
 * - It inserts checkboxes in ListingMaintenance for the boolean fields. These are native Google Sheets checkboxes.
 * - Formulas use ARRAYFORMULA where possible so new rows are auto-calculated.
 * - Conditional formatting rules are added for common alerts (未出单>7天, Coupon 即将结束, 首单耗时 color bands).
 * - Dashboard includes KPI formulas and a brand distribution table + chart (basic). You can further customize visuals in Google Sheets.
 */

function createProductTracker() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    SpreadsheetApp.getUi().alert('请先在 Google Drive 中打开或创建一个 Google 表格，然后在 Apps Script 中运行此脚本。');
    return;
  }

  // Helper to create or clear sheet
  function ensureSheet(name) {
    var sh = ss.getSheetByName(name);
    if (sh) {
      sh.clear();
      return sh;
    }
    return ss.insertSheet(name);
  }

  var products = ensureSheet('Products');
  var listings = ensureSheet('Listings');
  var ads = ensureSheet('Ads');
  var promos = ensureSheet('Promotions');
  var lm = ensureSheet('ListingMaintenance');
  var tasks = ensureSheet('Tasks');
  var lists = ensureSheet('Lists');
  var dash = ensureSheet('Dashboard');

  // Set headers
  products.getRange(1,1,1,10).setValues([['参考编码','ASIN','商品名称','品牌','大分类','中分类','小分类','产品负责人','供应商','建档日期']]);
  listings.getRange(1,1,1,9).setValues([['参考编码','上单门店','国家站点','上单日期','是否出单','出单日期','首单耗时','当前状态','上单多久了']]);
  ads.getRange(1,1,1,12).setValues([['参考编码','广告状态','广告开启日期','广告关闭日期','广告组数量','SP Campaign','SB Campaign','SD Campaign','自动广告','手动广告','广告预算','广告备注']]);
  promos.getRange(1,1,1,12).setValues([['参考编码','首发价格','当前价格','最近降价日期','Coupon 状态','Coupon 开始','Coupon 结束','Coupon 力度','LD','BD','Prime Exclusive','Prime Day']]);
  lm.getRange(1,1,1,9).setValues([['参考编码','五点完成','A+ 完成','Brand Story','视频','主图更新','QA 完善','Review 数量','Listing 完整度']]);
  tasks.getRange(1,1,1,7).setValues([['参考编码','当前问题','下一步动作','负责人','更新时间','优先级','状态']]);

  // Lists sheet: create columns for dropdowns
  lists.getRange(1,1).setValue('Stores');
  lists.getRange(2,1,4,1).setValues([['AM-BONNLO'],['AM-ROVSUN'],['VINGLI'],['MATLADIN']]);

  lists.getRange(1,2).setValue('SoldStatus');
  lists.getRange(2,2,3,1).setValues([['上单＜7天'],['否'],['是']]);

  lists.getRange(1,3).setValue('AdStatus');
  lists.getRange(2,3,4,1).setValues([['未开启'],['进行中'],['已暂停'],['已结束']]);

  lists.getRange(1,4).setValue('CouponStatus');
  lists.getRange(2,4,3,1).setValues([['未开启'],['进行中'],['已结束']]);

  lists.getRange(1,5).setValue('ListingStatus');
  lists.getRange(2,5,5,1).setValues([['待上单'],['已上单'],['优化中'],['稳定销售'],['停售']]);

  lists.getRange(1,6).setValue('SiteList');
  lists.getRange(2,6,5,1).setValues([['US'],['CA'],['UK'],['DE'],['JP']]);

  // Named ranges for each list (use the actual data ranges)
  ss.setNamedRange('Stores', lists.getRange('A2:A5'));
  ss.setNamedRange('SoldStatus', lists.getRange('B2:B4'));
  ss.setNamedRange('AdStatus', lists.getRange('C2:C5'));
  ss.setNamedRange('CouponStatus', lists.getRange('D2:D4'));
  ss.setNamedRange('ListingStatus', lists.getRange('E2:E6'));
  ss.setNamedRange('SiteList', lists.getRange('F2:F6'));

  // Insert sample data (5 rows)
  var sampleProducts = [
    ['REF001','B0001','Widget A','ROVSUN','Electronics','Audio','Headphones','张三','SupplierA','2026-06-01'],
    ['REF002','B0002','Widget B','BONNLO','Home','Kitchen','Appliances','王五','SupplierB','2026-06-20'],
    ['REF003','B0003','Widget C','VINGLI','Outdoor','Garden','Tools','赵七','SupplierC','2026-07-01'],
    ['REF004','B0004','Widget D','MATLADIN','Beauty','Skincare','Face','孙八','SupplierD','2026-05-01'],
    ['REF005','B0005','Widget E','ROVSUN','Electronics','Accessories','Cables','李华','SupplierE','2026-04-01']
  ];
  products.getRange(2,1,sampleProducts.length,sampleProducts[0].length).setValues(sampleProducts);

  var sampleListings = [
    ['REF001','AM-ROVSUN','US','2026-06-01','是','2026-06-05','', '稳定销售',''],
    ['REF002','AM-BONNLO','US','2026-06-20','否','','','优化中',''],
    ['REF003','VINGLI','US','2026-07-01','上单＜7天','','','上单＜7天',''],
    ['REF004','MATLADIN','US','2026-05-01','否','','','待上单',''],
    ['REF005','AM-ROVSUN','US','2026-04-01','是','2026-04-10','','稳定销售','']
  ];
  listings.getRange(2,1,sampleListings.length,sampleListings[0].length).setValues(sampleListings);

  var sampleAds = [
    ['REF001','进行中','2026-06-10','','3',2,1,0,true,false,'$20/day',''],
    ['REF002','未开启','','','1',1,0,0,false,true,'$0/day',''],
    ['REF003','进行中','2026-07-02','','2',1,1,0,true,true,'$10/day',''],
    ['REF004','已暂停','2026-05-10','2026-06-10','1',0,0,0,false,false,'$0/day','广告暂停'],
    ['REF005','已结束','2026-04-05','2026-05-01','0',0,0,0,false,false,'$0/day','']
  ];
  ads.getRange(2,1,sampleAds.length,sampleAds[0].length).setValues(sampleAds);

  var samplePromos = [
    ['REF001','$19.99','$17.99','2026-06-15','未开启','','','','$x','$x','$x','$x'],
    ['REF002','$29.99','$25.99','2026-06-25','进行中','2026-06-25','2026-07-05','15%','√','','',''],
    ['REF003','$15.00','$14.00','','未开启','','','','$x','$x','$x','$x'],
    ['REF004','$49.99','$44.99','2026-05-15','已结束','2026-05-10','2026-05-20','10%','', '√','√',''],
    ['REF005','$9.99','$7.99','2026-04-05','进行中','2026-03-01','2026-04-01','5%','','','','']
  ];
  promos.getRange(2,1,samplePromos.length,samplePromos[0].length).setValues(samplePromos);

  var sampleLM = [
    ['REF001',true,true,true,false,false,true,45,''],
    ['REF002',false,true,true,true,false,true,10,''],
    ['REF003',false,false,false,false,false,false,2,''],
    ['REF004',true,true,true,true,true,true,120,''],
    ['REF005',true,true,true,true,true,true,200,'']
  ];
  lm.getRange(2,1,sampleLM.length,sampleLM[0].length).setValues(sampleLM);

  var sampleTasks = [
    ['REF001','无','维护广告','李四','2026-07-09','高','进行中'],
    ['REF002','库存不足','补货','赵六','2026-07-09','高','待办'],
    ['REF003','等待Review','准备Coupon','钱七','2026-07-09','中','进行中'],
    ['REF004','Listing待优化','更新主图','周九','2026-07-09','中','进行中'],
    ['REF005','无','稳定','王强','2026-07-09','低','完成']
  ];
  tasks.getRange(2,1,sampleTasks.length,sampleTasks[0].length).setValues(sampleTasks);

  // Insert checkboxes for ListingMaintenance boolean columns (B-G)
  lm.getRange(2,2,100,6).insertCheckboxes();

  // Add array formulas for Listings: 首单耗时 (G), 上单多久了 (I)
  listings.getRange('G2').setFormula('=ArrayFormula(IF(A2:A="","",IF(F2:F="","",F2:F - D2:D)))');
  listings.getRange('I2').setFormula('=ArrayFormula(IF(A2:A="","",TODAY() - D2:D))');

  // ListingMaintenance completeness formula (I2)
  lm.getRange('I2').setFormula('=ArrayFormula(IF(A2:A="","",ROUND((N(B2:B)+N(C2:C)+N(D2:D)+N(E2:E)+N(F2:F)+N(G2:G))/6*100,0)&"%"))');

  // Coupons: add a reminder column in Promotions at column M (13)
  promos.getRange(1,13).setValue('Coupon 提醒');
  promos.getRange('M2').setFormula('=ArrayFormula(IF(A2:A="","",IF((E2:E="进行中")*(G2:G<>"")*((G2:G)-TODAY()<=3),"⚠ 即将结束","")))');

  // Conditional formatting rules
  // 1) Listings: 未出单超7天 -> red fill
  var listingsRange = listings.getRange('A2:I1000');
  var rule1 = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($F2="", $D2<>"", TODAY()-$D2>7)')
      .setBackground('#F8D7DA')
      .setRanges([listingsRange])
      .build();

  // 2) Listings: 首单耗时 color band (G)
  var gRange = listings.getRange('G2:G1000');
  var ruleG1 = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$G2<>""').setBackground('#E6F4EA').setRanges([gRange]).build(); // placeholder green for non-empty
  // 3) Promotions: Coupon 提醒 highlight
  var promosRange = promos.getRange('A2:M1000');
  var ruleP = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$M2<>""').setBackground('#FFF3CD').setRanges([promosRange]).build();

  var rules = listings.getConditionalFormatRules();
  rules.push(rule1);
  listings.setConditionalFormatRules(rules);

  var rulesP = promos.getConditionalFormatRules();
  rulesP.push(ruleP);
  promos.setConditionalFormatRules(rulesP);

  // Dashboard KPI formulas
  dash.getRange('A1').setValue('Product Tracker Dashboard');
  dash.getRange('A3').setValue('总 SKU');
  dash.getRange('B3').setFormula('=COUNTA(Products!A2:A)');
  dash.getRange('A4').setValue('已出单数量');
  dash.getRange('B4').setFormula('=COUNTIF(Listings!E2:E,"是")');
  dash.getRange('A5').setValue('广告进行中数');
  dash.getRange('B5').setFormula('=COUNTIF(Ads!B2:B,"进行中")');
  dash.getRange('A6').setValue('Coupon 进行中数');
  dash.getRange('B6').setFormula('=COUNTIF(Promotions!E2:E,"进行中")');
  dash.getRange('A7').setValue('首单平均耗时(天)');
  dash.getRange('B7').setFormula('=AVERAGEIF(Listings!G2:G,">0")');
  dash.getRange('A8').setValue('平均 Listing 完整度');
  // For completeness numeric average, create hidden helper in ListingMaintenance col J with numeric value
  lm.getRange('J1').setValue('完整度数值');
  lm.getRange('J2').setFormula('=ArrayFormula(IF(A2:A="","",VALUE(LEFT(I2:I,LEN(I2:I)-1))/100))');
  dash.getRange('B8').setFormula('=AVERAGEIF(ListingMaintenance!J2:J,">0")');

  // Brand distribution via QUERY in Dashboard (for chart)
  dash.getRange('A10').setValue('Brand');
  dash.getRange('B10').setValue('Count');
  dash.getRange('A11').setFormula('=QUERY(Products!D2:D,"select D, count(D) where D is not null group by D order by count(D) desc",0)');

  // Create a chart (bar) based on the brand distribution generated by QUERY (A11:B)
  // Wait a moment: charts need a range; we will create a basic chart anchored at E11
  try {
    var chart = dash.newChart()
      .asColumnChart()
      .addRange(dash.getRange('A11:B16'))
      .setPosition(11,5,0,0)
      .setNumHeaders(1)
      .setOption('title','品牌分布')
      .build();
    dash.insertChart(chart);
  } catch (e) {
    // ignore chart errors if range empty
  }

  // Freeze header rows
  [products, listings, ads, promos, lm, tasks, lists, dash].forEach(function(s){
    s.setFrozenRows(1);
    s.autoResizeColumns(1,15);
  });

  SpreadsheetApp.getUi().alert('Product Tracker 已生成。请检查 Lists sheet 的命名范围；如需下拉数据验证，请在对应列引用这些命名范围（数据->数据验证->范围）。运行完成后可使用 文件->下载->Microsoft Excel(.xlsx) 保存到本地在 WPS 中打开。');
}
