async function main(workbook: ExcelScript.Workbook) {
  // Call the custom REST API with a POST request
  const response = await fetch('https://testservercomp.jcampbell2.repl.co/funnypost', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      // Replace this object with the data you want to send
      userId: 1,
      id: 1,
      title: "string",
      completed: true
    })
});

  const data: Data = await response.json();

  const rows: (string | boolean | number)[][] = [];

  rows.push([data.userId, data.id, data.title, data.completed]);

  const sheet = workbook.getActiveWorksheet();
  sheet.getRange('A1:D1').setValues([["User ID", "ID", "Title", "Completed"]]);

  const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
  range.setValues(rows);
}

// An interface matching the returned JSON for the data.
interface Data {
  userId: number,
  id: number,
  title: string,
  completed: boolean
}
