const express = require("express");
const mysql = require("./database/config");
const { generateExcelFile, generateExcelFilePickup } = require("./excel");

const app = express();
const port = 3000;

app.get("/delivery-information", async (req, res) => {
  try {
    // query users table
    const usersQuery = "SELECT * FROM completeddelivery";
    const [usersRows] = await mysql.promise().query(usersQuery);

    // iterate through users
    const rows = [];

    for (const row_data of usersRows) {
      const cartStart = row_data.CartStart;
      const cartEnd = row_data.CartEnd;

      // query cart table based on cartStart and cartEnd
      const cartQuery = `SELECT * FROM cart WHERE cartId BETWEEN '${cartStart}' AND '${cartEnd}'`;
      const [cartRows] = await mysql.promise().query(cartQuery);

      // iterate through cart items
      for (const cartRow of cartRows) {
        const productId = cartRow.ProductId;
        const quantity = cartRow.Quantity;
        const unitPrice = cartRow.price;
        const total = cartRow.Amount;

        // query product table based on productId
        const productQuery = `SELECT * FROM products WHERE productId=${productId}`;
        const [productRows] = await mysql.promise().query(productQuery);
        const productRow = productRows[0];

        const productName = productRow.Title;
        const productPrice = productRow.price;
        const productType = productRow.Description;
        const productRoast = productRow.Roast;
        const productSize = productRow.size;

        // add data to current row
        rows.push({
          User: row_data.FirstName,
          LastName: row_data.LastName,
          "Phone Number": row_data.PhoneNumber,
          TinName: row_data.TinName,
          TinNumber: row_data.TinNumber,
          Product: productName,
          productType: productType,
          Size: productSize,
          Roast: productRoast,
          Quantity: quantity,
          "Unit Price": productPrice,
          Total: total,
          ShopLocation: row_data.ShopLocation,
          OrderNumber: row_data.OrderNumber,
          CompletedDate: row_data.CompletedDate.toString(),
        });
      }
    }
    // console.log(rows);
    // generate Excel file
    const excelFile = generateExcelFile(rows);

    generateExcelFile(rows)
      .then((rresultes) => {
        // do something with the buffer, e.g. write it to a file
        res.setHeader("Content-Type", "application/vnd.ms-excel");
        res.setHeader(
          "Content-Disposition",
          "attachment; filename=cart_information.xlsx"
        );
        res.send(rresultes);
      })
      .catch((err) => {
        console.error(err);
      });

    // send Excel file to client
  } catch (error) {
    console.error(error);
    res.status(500).send("Internal Server Error");
  }
});

app.get("/pickedup-information", async (req, res) => {
  try {
    // query users table
    const usersQuery = "SELECT * FROM pickupcompleted";
    const [usersRows] = await mysql.promise().query(usersQuery);

    // iterate through users
    const rows = [];

    for (const row_data of usersRows) {
      const cartStart = row_data.CartStart;
      const cartEnd = row_data.CartEnd;

      // query cart table based on cartStart and cartEnd
      const cartQuery = `SELECT * FROM cart WHERE cartId BETWEEN '${cartStart}' AND '${cartEnd}'`;
      const [cartRows] = await mysql.promise().query(cartQuery);

      // iterate through cart items
      for (const cartRow of cartRows) {
        const productId = cartRow.ProductId;
        const quantity = cartRow.Quantity;
        const unitPrice = cartRow.price;
        const total = cartRow.Amount;

        // query product table based on productId
        const productQuery = `SELECT * FROM products WHERE productId=${productId}`;
        const [productRows] = await mysql.promise().query(productQuery);
        const productRow = productRows[0];

        const productName = productRow.Title;
        const productPrice = productRow.price;
        const productType = productRow.Description;
        const productRoast = productRow.Roast;
        const productSize = productRow.size;

        // add data to current row
        rows.push({
          User: row_data.FirstName,
          LastName: row_data.LastName,
          "Phone Number": row_data.PhonNumber,
          Product: productName,
          productType: productType,
          Size: productSize,
          Roast: productRoast,
          Quantity: quantity,
          "Unit Price": productPrice,
          Total: total,
          ShopLocation: row_data.Shop,
          OrderNumber: row_data.OrderNumber,
        });
      }
    }
    // console.log(rows);
    // generate Excel file
    const excelFile = generateExcelFilePickup(rows);

    generateExcelFilePickup(rows)
      .then((rresultes) => {
        // do something with the buffer, e.g. write it to a file
        res.setHeader("Content-Type", "application/vnd.ms-excel");
        res.setHeader(
          "Content-Disposition",
          "attachment; filename=cart_information.xlsx"
        );
        res.send(rresultes);
      })
      .catch((err) => {
        console.error(err);
      });

    // send Excel file to client
  } catch (error) {
    console.error(error);
    res.status(500).send("Internal Server Error");
  }
});

app.listen(port, () => {
  console.log("Server listening at http://localhost:", port);
});
