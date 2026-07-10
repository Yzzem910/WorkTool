/**
 * Office Script: generate_product_tracker_office_script.ts
 *
 * 说明：
 * - 在 Excel for Web (Office Scripts) 的 Automate -> New Script 中粘贴并运行此脚本。
 * - 脚本会在当前工作簿中新建多个 sheet：Products, Listings, Ads, Promotions, ListingMaintenance, Tasks, Lists, Dashboard
 * - 会建立表格（Table_*），写入表头与 5 行示例数据，以及基础计算公式（使用结构化引用时，部分公式会在表格创建后以 A1 形式插入以保持兼容）。
 * - 数据验证（下拉）、复选框（控件）和透视表/图表在 Office Scripts 中支持有限；脚本会放置提示并写入 Lists 里供你手动在 Excel 客户端通过“数据 > 数据验证”引用。
 *
 * 使用步骤：
 * 1) 在 Excel for Web 打开一个空白工作簿。
 * 2) Automate -> New Script，替换默认代码，粘贴并保存本脚本。
 * 3) 运行脚本（Run）。脚本完成后另存为 .xlsx（File -> Save As -> 下载）即可在本地 Excel 打开并继续美化（添加复选控件或切片器等）。
 */
function main(workbook: ExcelScript.Workbook) {
  // Utility to add sheet if not exists
  function addSheet(name: string) {
    let sh = workbook.getWorksheet(name);
    if (sh) {
      // clear content
      sh.getRange().clear(ExcelScript.ClearApplyTo.contents);
      return sh;
    }
    return workbook.addWorksheet(name);
  }

  // 1) Products sheet
  let products = addSheet("Products");
  let productsHeaders = [
    [
      "参考编码","ASIN","商品名称","品牌","大分类","中分类","小分类","产品负责人","供应商","建档日期"
    ]
  ];
  products.getRangeByIndexes(0,0,1,productsHeaders[0].length).setValues(productsHeaders);

  // 2) Listings sheet
  let listings = addSheet("Listings");
  let listingsHeaders = [
    [
      "参考编码","上单门店","国家站点","上单日期","是否出单","出单日期","首单耗时","当前状态","上单多久了"
    ]
  ];
  listings.getRangeByIndexes(0,0,1,listingsHeaders[0].length).setValues(listingsHeaders);

  // 3) Ads sheet
  let ads = addSheet("Ads");
  let adsHeaders = [
    [
      "参考编码","广告状态","广告开启日期","广告关闭日期","广告组数量","SP Campaign","SB Campaign","SD Campaign","自动广告","手动广告","广告预算","广告备注"
    ]
  ];
  ads.getRangeByIndexes(0,0,1,adsHeaders[0].length).setValues(adsHeaders);

  // 4) Promotions sheet
  let promos = addSheet("Promotions");
  let promosHeaders = [
    [
      "参考编码","首发价格","当前价格","最近降价日期","Coupon 状态","Coupon 开始","Coupon 结束","Coupon 力度","LD","BD","Prime Exclusive","Prime Day"
    ]
  ];
  promos.getRangeByIndexes(0,0,1,promosHeaders[0].length).setValues(promosHeaders);

  // 5) ListingMaintenance sheet
  let lm = addSheet("ListingMaintenance");
  let lmHeaders = [
    [
      "参考编码","五点完成","A+ 完成","Brand Story","视频","主图更新","QA 完善","Review 数量","Listing 完整度"
    ]
  ];
  lm.getRangeByIndexes(0,0,1,lmHeaders[0].length).setValues(lmHeaders);

  // 6) Tasks sheet
  let tasks = addSheet("Tasks");
  let tasksHeaders = [["参考编码","当前问题","下一步动作","负责人","更新时间","优先级","状态"]];
  tasks.getRangeByIndexes(0,0,1,tasksHeaders[0].length).setValues(tasksHeaders);

  // 7) Lists sheet (下拉项集中区)
  let lists = addSheet("Lists");
  // We'll write lists vertically with a header row name
  lists.getRange("A1").setValue("Stores");
  lists.getRange("A2:A5").setValues([["AM-BONNLO"],["AM-ROVSUN"],["VINGLI"],["MATLADIN"]]);

  lists.getRange("B1").setValue("SoldStatus");
  lists.getRange("B2:B4").setValues([["上单＜7天"],["否"],["是"]]);

  lists.getRange("C1").setValue("AdStatus");
  lists.getRange("C2:C5").setValues([["未开启"],["进行中"],["已暂停"],["已结束"]]);

  lists.getRange("D1").setValue("CouponStatus");
  lists.getRange("D2:D4").setValues([["未开启"],["进行中"],["已结束"]]);

  lists.getRange("E1").setValue("ListingStatus");
  lists.getRange("E2:E6").setValues([["待上单"],["已上单"],["优化中"],["稳定销售"],["停售"]]);

  lists.getRange("F1").setValue("SiteList");
  lists.getRange("F2:F6").setValues([["US"],["CA"],["UK"],["DE"],["JP"]]);

  // 8) Dashboard sheet (placeholder)
  let dash = addSheet("Dashboard");
  dash.getRange("A1").setValue("Product Tracker Dashboard (请在 Excel 中刷新/美化图表)");

  // Create sample data (5 rows) - use the company choosed primary key: 参考编码
  let sampleProducts = [
    ["REF001","B0001","Widget A","ROVSUN","Electronics","Audio","Headphones","张三","SupplierA","2026-06-01"],
    ["REF002","B0002","Widget B","BONNLO","Home","Kitchen","Appliances","王五","SupplierB","2026-06-20"],
    ["REF003","B0003","Widget C","VINGLI","Outdoor","Garden","Tools","赵七","SupplierC","2026-07-01"],
    ["REF004","B0004","Widget D","MATLADIN","Beauty","Skincare","Face","孙八","SupplierD","2026-05-01"],
    ["REF005","B0005","Widget E","ROVSUN","Electronics","Accessories","Cables","李华","SupplierE","2026-04-01"]
  ];

  products.getRangeByIndexes(1,0,sampleProducts.length,sampleProducts[0].length).setValues(sampleProducts);

  // Sample Listings rows (matching by 参考编码)
  let sampleListings = [
    ["REF001","AM-ROVSUN","US","2026-06-01","是","2026-06-05","=F2-E2","稳定销售","=TODAY()-E2"],
    ["REF002","AM-BONNLO","US","2026-06-20","否","","","优化中","=IF(E3=="","",TODAY()-E3)"],
    ["REF003","VINGLI","US","2026-07-01","上单＜7天","","","上单＜7天","=TODAY()-E4"],
    ["REF004","MATLADIN","US","2026-05-01","否","","","待上单","=TODAY()-E5"],
    ["REF005","AM-ROVSUN","US","2026-04-01","是","2026-04-10","=F6-E6","稳定销售","=TODAY()-E6"]
  ];
  listings.getRangeByIndexes(1,0,sampleListings.length,sampleListings[0].length).setValues(sampleListings);

  // Sample Ads
  let sampleAds = [
    ["REF001","进行中","2026-06-10","","3",2,1,0,true,false,"$20/day",""],
    ["REF002","未开启","","","1",1,0,0,false,true,"$0/day",""],
    ["REF003","进行中","2026-07-02","","2",1,1,0,true,true,"$10/day",""],
    ["REF004","已暂停","2026-05-10","2026-06-10","1",0,0,0,false,false,"$0/day","广告暂停"],
    ["REF005","已结束","2026-04-05","2026-05-01","0",0,0,0,false,false,"$0/day",""]
  ];
  ads.getRangeByIndexes(1,0,sampleAds.length,sampleAds[0].length).setValues(sampleAds);

  // Sample Promotions
  let samplePromos = [
    ["REF001","$19.99","$17.99","2026-06-15","未开启","","","","×","×","×","×"],
    ["REF002","$29.99","$25.99","2026-06-25","进行中","2026-06-25","2026-07-05","15%","√","×","×","×"],
    ["REF003","$15.00","$14.00","","未开启","","","","×","×","×","×"],
    ["REF004","$49.99","$44.99","2026-05-15","已结束","2026-05-10","2026-05-20","10%","×","√","√","×"],
    ["REF005","$9.99","$7.99","2026-04-05","进行中","2026-03-01","2026-04-01","5%","×","×","×","×"]
  ];
  promos.getRangeByIndexes(1,0,samplePromos.length,samplePromos[0].length).setValues(samplePromos);

  // Sample ListingMaintenance
  let sampleLM = [
    ["REF001",true,true,true,false,false,true,45,"=ROUND((IF(B2,1,0)+IF(C2,1,0)+IF(D2,1,0)+IF(E2,1,0)+IF(F2,1,0)+IF(G2,1,0))/6*100,0)&"%""],
    ["REF002",false,true,true,true,false,true,10,"=ROUND((IF(B3,1,0)+IF(C3,1,0)+IF(D3,1,0)+IF(E3,1,0)+IF(F3,1,0)+IF(G3,1,0))/6*100,0)&"%""],
    ["REF003",false,false,false,false,false,false,2,"=ROUND((IF(B4,1,0)+IF(C4,1,0)+IF(D4,1,0)+IF(E4,1,0)+IF(F4,1,0)+IF(G4,1,0))/6*100,0)&"%""],
    ["REF004",true,true,true,true,true,true,120,"=ROUND((IF(B5,1,0)+IF(C5,1,0)+IF(D5,1,0)+IF(E5,1,0)+IF(F5,1,0)+IF(G5,1,0))/6*100,0)&"%""],
    ["REF005",true,true,true,true,true,true,200,"=ROUND((IF(B6,1,0)+IF(C6,1,0)+IF(D6,1,0)+IF(E6,1,0)+IF(F6,1,0)+IF(G6,1,0))/6*100,0)&"%""],
  ];
  // Note: above ListingMaintenance formulas are left as string placeholders; users may adjust within Excel if needed.
  lm.getRangeByIndexes(1,0,sampleLM.length,8).setValues(sampleLM.map(r => r.slice(0,8)));

  // Sample Tasks
  let sampleTasks = [
    ["REF001","无","维护广告","李四","2026-07-09","高","进行中"],
    ["REF002","库存不足","补货","赵六","2026-07-09","高","待办"],
    ["REF003","等待Review","准备Coupon","钱七","2026-07-09","中","进行中"],
    ["REF004","Listing待优化","更新主图","周九","2026-07-09","中","进行中"],
    ["REF005","无","稳定","王强","2026-07-09","低","完成"]
  ];
  tasks.getRangeByIndexes(1,0,sampleTasks.length,sampleTasks[0].length).setValues(sampleTasks);

  // Convert ranges to tables for better structure
  function createTableIfNotExists(sheet: ExcelScript.Worksheet, headerCount: number, sampleRowCount: number, tableName: string) {
    let usedRange = sheet.getUsedRange();
    if (!usedRange) return null;
    let rowCount = usedRange.getRowCount();
    let colCount = usedRange.getColumnCount();
    // Table range: header + sampleRowCount rows
    let tblRange = sheet.getRangeByIndexes(0,0,rowCount,colCount);
    try {
      let tbl = sheet.addTable(tblRange.getAddress(), true);
      tbl.setName(tableName);
      // style header fill color by tableName (simple palette)
      let header = tbl.getHeaderRowRange();
      let color = "#DCEEF8"; // default light blue
      if (tableName.indexOf("Listings")>=0) color = "#E6F4EA"; // light green
      if (tableName.indexOf("Ads")>=0) color = "#FFEFD5"; // light orange
      if (tableName.indexOf("Promotions")>=0) color = "#FFF0F6"; // light pink
      if (tableName.indexOf("ListingMaintenance")>=0) color = "#F3E8FF"; // light purple
      if (tableName.indexOf("Tasks")>=0) color = "#F2F2F2"; // grey
      header.getFormat().getFill().setColor(color);
      return tbl;
    } catch (e) {
      // ignore
      return null;
    }
  }

  createTableIfNotExists(products, productsHeaders[0].length, sampleProducts.length, "Table_Products");
  createTableIfNotExists(listings, listingsHeaders[0].length, sampleListings.length, "Table_Listings");
  createTableIfNotExists(ads, adsHeaders[0].length, sampleAds.length, "Table_Ads");
  createTableIfNotExists(promos, promosHeaders[0].length, samplePromos.length, "Table_Promotions");
  createTableIfNotExists(lm, lmHeaders[0].length, sampleLM.length, "Table_ListingMaintenance");
  createTableIfNotExists(tasks, tasksHeaders[0].length, sampleTasks.length, "Table_Tasks");

  // Freeze top rows for each sheet
  [products, listings, ads, promos, lm, tasks, lists, dash].forEach(s => s.freezePanes.freezeRows(1));

  // Autofit columns for readability
  [products, listings, ads, promos, lm, tasks, lists, dash].forEach(s => s.getUsedRange()?.getFormat().autofitColumns());

  // Instruction note in Dashboard
  dash.getRange("A3").setValue("说明：请在 Excel 客户端内为 Listings/Ads/Promotions 的指定列添加数据验证（引用 Lists sheet 中对应的选项），以启用下拉菜单。若需复选样式，请用表单控件或在列中使用 TRUE/FALSE 并将单元格格式设置为勾选样式。\n在桌面 Excel 中你可以进一步添加透视表/切片器和图表以完成 Dashboard。脚本已生成表格和示例公式占位。\n如果需要，我可以随后将一个 VBA 脚本或更复杂的脚本添加到仓库以自动化更多交互元素（例如复选框控件、切片器和邮件提醒）。");

  // Final note
  dash.getRange("A10").setValue("脚本完成：请保存工作簿为 .xlsx 并在本地打开以启用完整 Excel 功能（复选控件、宏、切片器等）。");
}
