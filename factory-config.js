window.FACTORY_CONFIG = Object.freeze({
  identity: Object.freeze({
    plantCode: "GS1",
    plantName: "Nhà máy Goldsun Hà Nội",
    pageTitle: "GS1 · Nhà máy Goldsun Hà Nội · Smart Factory",
    dashboardTitle:
      "GS1 · Nhà máy Goldsun Hà Nội · Dashboard sản xuất từ dữ liệu Lệnh thao tác",
    otherSegmentLabel: "99_AF chưa ánh xạ",
    storageNamespace: "gsp.factory.gs1",
    taskSourceDescription:
      "Chưa cấu hình 00_Task_Schedule riêng cho GS1. Mô-đun công việc được khóa để không dùng nhầm dữ liệu nhà máy khác."
  }),
  production: Object.freeze({
    url:
      "https://docs.google.com/spreadsheets/d/1c3_CSdnh9sxALEt-6FKyrd-4-J1K_Yhl",
    fileId: "1c3_CSdnh9sxALEt-6FKyrd-4-J1K_Yhl",
    fileName: "P3_Tong_Hop_LTT_2507-HN.xlsx",
    sheetName: "P3.Tổng hợp lệnh thao tác",
    range: "A9:CT",
    manifestUrl: "./data/manifest.json"
  }),
  dashboard: Object.freeze({
    segmentMode: "gs1_parent_line_af"
  }),
  tasks: Object.freeze({
    enabled: false,
    url: "",
    sheetName: "00_Task_Schedule",
    range: "A3:AA"
  })
});
