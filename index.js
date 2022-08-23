var Excel = require('exceljs');
var workbook = new Excel.Workbook();
const fs = require('fs');
var nodeoutlook = require('nodejs-nodemailer-outlook');

workbook.xlsx.readFile('./august.xlsx').then(function () {
  ws = workbook.getWorksheet('Sheet1');
  const customer = new Object();
  const customers = [];
  customer.customerType = [];
  customer.customerEmail = [];
  customer.orderDate = [];
  let nonQuotaCustomerEmails = [];

  ws.getColumn(9).eachCell((cell, rn) => {
    customer.customerType.push(cell.value);
  });
  ws.getColumn(8).eachCell((cell, rn) => {
    customer.customerEmail.push(cell.value);
  });
  ws.getColumn(14).eachCell((cell, rn) => {
    customer.orderDate.push(cell.value);
  });

  customer.customerEmail.shift();
  customer.customerType.shift();
  customer.orderDate.shift();

  for (i in customer.customerType) {
    if (customer.customerType[i] == 'Q3') {
      nonQuotaCustomerEmails.push(customer.customerEmail[i]);
      let inputDate = customer.orderDate[i];

      let dateInteger = new Date(inputDate).getTime();

      function getDuration(milli) {
        let today = Date.now();

        let todayMinutes = Math.floor(today / 60000);
        let todayHours = Math.round(todayMinutes / 60);
        let todayDays = Math.round(todayHours / 24);

        let minutes = Math.floor(milli / 60000);
        let hours = Math.round(minutes / 60);
        let days = Math.round(hours / 24);

        let date = new Date(days * 8.64e7).toISOString();
        return (
          (days && {
            daysSinceEpoch: days,
            unit: 'days',
            purchaseDate: date,
            daysDifference: todayDays - days,
          }) ||
          (hours && { daysSinceEpoch: hours, unit: 'hours' }) || {
            daysSinceEpoch: minutes,
            unit: 'minutes',
          }
        );
      }

      customers.push({
        type: customer.customerType[i],
        email: customer.customerEmail[i],
        orderDate: getDuration(dateInteger),
      });
    }
  }

  const jsonString = JSON.stringify(customers, null, 2);

  fs.writeFile('./Q3-Customers.json', jsonString, (err) => {
    if (err) {
      console.log('Error writing file', err);
    } else {
      console.log('Successfully wrote file');
      let userData = fs.readFileSync('./Q3-Customers.json');
      let userDataEmailed = fs.readFileSync('./Q3-Customers-Emailed.json');
      let customerRead = JSON.parse(userData);
      let customerEmailedRead = JSON.parse(userDataEmailed);
      let withinTwoWeeks = [];

      for (i in customerRead) {
        if (customerRead[i].orderDate.daysDifference < 14)
          withinTwoWeeks.push(customerRead[i]);
      }

      // Helper section for determining differences between arrays of objects
      const isSameUser = (a, b) => a.email === b.email;

      const onlyInLeft = (left, right, compareFunction) =>
        left.filter(
          (leftValue) =>
            !right.some((rightValue) => compareFunction(leftValue, rightValue))
        );

      const onlyInA = onlyInLeft(
        withinTwoWeeks,
        customerEmailedRead,
        isSameUser
      );
      const onlyInB = onlyInLeft(
        customerEmailedRead,
        withinTwoWeeks,
        isSameUser
      );

      const result = [...onlyInA, ...onlyInB];
      console.log(onlyInA);

      for (i in onlyInA) {
        customerEmailedRead.push(onlyInA[i]);
      }

      let updatedEmailed = JSON.stringify(customerEmailedRead, null, 2);

      fs.writeFile('./Q3-Customers-Emailed.json', updatedEmailed, (err) => {
        if (err) {
          console.log('Error writing file', err);
        } else {
          console.log('Successfully wrote file');
        }
      });
    }
  });
});
