/* eslint-disable no-console */
const fs = require('fs');
// FIXME: Incase you have the npm package
const HTMLtoDOCX = require('html-to-docx');

const filePath = 'test.docx';

const htmlString = `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sample Page</title>

</head>
<body>
    <h1>Heading 1</h1>
    <p style="color: red;"><em>Lorem ipsum dolor sit amet</em>, consectetur adipiscing elit. <strong>Integer scelerisque</strong> orci vel arcu gravida sollicitudin.</p>
    
    <h2>Heading 2</h2>
    <p style="color: blue;">Sed non <strong>efficitur diam</strong>. Vivamus <em>convallis dolor</em> at eros scelerisque, in tincidunt turpis consequat.</p>
    
    <h3>Heading 3</h3>
    <p style="color: green;">Proin sit amet <em>placerat odio</em>. Fusce <strong>malesuada feugiat</strong> ligula, nec rutrum libero consectetur nec.</p>
    
    <h4>Heading 4</h4>
    <p>In vitae enim sit amet <em>arcu placerat posuere</em> eget vel quam. Sed malesuada, <strong>nisl in vestibulum</strong> accumsan, libero nulla ornare nisi, vel varius mi enim sed elit.</p>
    
    <table border="1">
        <tr>
            <th>Column 1</th>
            <th>Column 2</th>
        </tr>
        <tr>
            <td><em>Row 1, Cell 1</em></td>
            <td><strong>Row 1, Cell 2</strong></td>
        </tr>
        <tr>
            <td><strong>Row 2, Cell 1</strong></td>
            <td><em>Row 2, Cell 2</em></td>
        </tr>
    </table>
    <h2> Table example</h2>
    <table>
        <thead>
            <tr>
                <th colspan="2">Employee Details</th>
                <th rowspan="2">Department</th>
                <th colspan="2">Salary</th>
            </tr>
            <tr>
                <th>Name</th>
                <th>Age</th>
                <th>Basic</th>
                <th>Bonus</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>John Doe</td>
                <td>30</td>
                <td>Engineering</td>
                <td>$5000</td>
                <td>$1000</td>
            </tr>
            <tr>
                <td>Mary Smith</td>
                <td>35</td>
                <td>Marketing</td>
                <td>$4500</td>
                <td>$800</td>
            </tr>
            <tr>
                <td>David Johnson</td>
                <td>40</td>
                <td>Finance</td>
                <td>$6000</td>
                <td>$1200</td>
            </tr>
        </tbody>
    </table>
    
    <h2>Images</h2>
    <div>
        <p><strong>Image 1:</strong> Description of the image</p>
        <img src="1.jpg" width="700" height="500" alt="Image 1">
    </div>
    
    
    <div>
        <p><strong>Image 1:</strong> Description of the image</p>
        <img src="1.jpg" width="700" height="500" alt="Image 1">
    </div>
</body>
</html>

`;

(async () => {
  const fileBuffer = await HTMLtoDOCX(htmlString, null, {
    table: { row: { cantSplit: true } },
    footer: true,
    pageNumber: true,
  });

  fs.writeFile(filePath, fileBuffer, (error) => {
    if (error) {
      console.log('Docx file creation failed');
      console.log(error)
      return;
    }
    console.log('Docx file created successfully');
  });
})();
