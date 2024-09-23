<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Export HTML to Excel</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.0/xlsx.full.min.js"></script>
</head>
<body>

    <button id="exportBtn">Export to Excel</button>

    <!-- Your HTML table here -->
    <table class="ab_report">
        <tr>
            <th>ITEM NO.</th>
            <th>ITEM & DESCRIPTION</th>
            <th>QTY</th>
            <th>Estimate Cost</th>
            <th>Unit Price</th>
            <th>Total Price</th>
        </tr>
        <tr>
            <td>1</td>
            <td>1 meal, 2 snacks</td>
            <td>40</td>
            <td>450</td>
            <td>500</td>
            <td>20,000.00</td>
        </tr>
        <!-- Add more rows as needed -->
    </table>

    <script>
        document.getElementById('exportBtn').addEventListener('click', function() {
            var wb = XLSX.utils.book_new();
            var table = document.querySelector('table');
            var ws = XLSX.utils.table_to_sheet(table);
            XLSX.utils.book_append_sheet(wb, ws, 'Abstract of Quotation');
            XLSX.writeFile(wb, 'Abstract_of_Quotation.xlsx');
        });
    </script>

</body>
</html>