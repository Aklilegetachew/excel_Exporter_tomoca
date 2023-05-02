const ExcelJS = require("exceljs");

const generateExcelFile = async (rows) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Cart Information");

  // set column headers
  worksheet.columns = [
    { header: "First Name", key: "User", width: 20 },
    { header: "Last Name", key: "LastName", width: 20 },
    { header: "Phone Number", key: "Phone Number", width: 20 },
    { header: "TinName", key: "TinName", width: 20 },
    { header: "TinNumber", key: "TinNumber", width: 20 },
    { header: "Shop Location", key: "ShopLocation", width: 15 },
    { header: "Completed Date", key: "CompletedDate", width: 15 },
    { header: "Order Number", key: "OrderNumber", width: 15 },
    { header: "Product", key: "Product", width: 30 },
    { header: "Size", key: "Size", width: 15 },
    { header: "Type", key: "productType", width: 15 },

    { header: "Roast", key: "Roast", width: 15 },
    { header: "Quantity", key: "Quantity", width: 10 },
    { header: "Unit Price", key: "Unit Price", width: 15 },
    { header: "Total", key: "Total", width: 150 },
  ];

  // add data to rows
  // add data to rows
  rows.forEach((row) => {
    worksheet.addRow(row);
  });

  // return buffer
  const buffer = await workbook.xlsx.writeBuffer();
  return buffer;
};

const generateExcelFilePickup = async (rows) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Cart Information");

  // set column headers
  worksheet.columns = [
    { header: "First Name", key: "User", width: 20 },
    { header: "Last Name", key: "LastName", width: 20 },
    { header: "Phone Number", key: "Phone Number", width: 20 },

    { header: "Shop Location", key: "ShopLocation", width: 15 },

    { header: "Order Number", key: "OrderNumber", width: 15 },
    { header: "Product", key: "Product", width: 30 },
    { header: "Size", key: "Size", width: 15 },
    { header: "Type", key: "productType", width: 15 },

    { header: "Roast", key: "Roast", width: 15 },
    { header: "Quantity", key: "Quantity", width: 10 },
    { header: "Unit Price", key: "Unit Price", width: 15 },
    { header: "Total", key: "Total", width: 150 },
  ];

  // add data to rows
  // add data to rows
  rows.forEach((row) => {
    worksheet.addRow(row);
  });

  // return buffer
  const buffer = await workbook.xlsx.writeBuffer();
  return buffer;
};

module.exports = { generateExcelFile, generateExcelFilePickup };
