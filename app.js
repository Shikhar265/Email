const express = require('express');
const bodyParser = require('body-parser');
const path = require('path');
const xlsx = require('xlsx');
const app = express();
let formFields = {};
let formTitle = '';

app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// Index route (form title entry page)
app.get('/', (req, res) => {
  res.render('index');
});

// Process form title and redirect to form field creation
app.post('/create-form', (req, res) => {
  formTitle = req.body.formTitle;
  res.redirect('/create-form');
});

// Form field creation page (teacher's side)
app.get('/create-form', (req, res) => {
  res.render('create-form', { formTitle });
});

// Handle form fields input and generate the form URL
app.post('/save-fields', (req, res) => {
  const fields = req.body.fields.split(',').map(field => field.trim());
  formFields[formTitle] = fields;
  const formUrl = `http://localhost:3000/fill-form/${encodeURIComponent(formTitle)}`;
  res.send(`Form created! Share this URL with students: <a href="${formUrl}">${formUrl}</a>`);
});

// Student form page (student's side)
app.get('/fill-form/:formTitle', (req, res) => {
  const title = req.params.formTitle;
  const fields = formFields[title] || [];
  res.render('fill-form', { formTitle: title, formFields: fields });
});

// Handle form submission and save to Excel
app.post('/submit-form/:formTitle', (req, res) => {
  const title = req.params.formTitle;
  const data = req.body;
  const workbook = xlsx.utils.book_new();
  const worksheet = xlsx.utils.json_to_sheet([data]);
  xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  const filePath = path.join(__dirname, `${title}.xlsx`);
  xlsx.writeFileSync(workbook, filePath);
  res.send('Form submitted successfully and saved to Excel!');
});

// Set EJS as the view engine
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

app.listen(3000, () => {
  console.log('Server running at http://localhost:3000');
});
