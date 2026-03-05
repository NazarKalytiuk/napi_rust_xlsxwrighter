const { Workbook, Note } = require('@nazarkalytiuk/rusc-xlsx');

console.log('Creating Excel file with cell notes...');

// Create a new workbook
const workbook = new Workbook();

// Add a worksheet
const worksheet = workbook.addWorksheet('Party Planning');

// Set the default author for all notes in this worksheet
worksheet.setDefaultNoteAuthor('Clarissa Dalloway');

// Widen the first column for clarity
worksheet.setColumnWidth(0, 16);

// Write some data
const partyItems = ['Invitations', 'Doors', 'Flowers', 'Champagne', 'Menu', 'Peter'];

partyItems.forEach((item, index) => {
    worksheet.writeString(index, 0, item);
});

// Create a basic note
const note1 = new Note('I will get the flowers myself');
note1.setAuthor('Clarissa Dalloway');
worksheet.insertNote(2, 0, note1); // Add note to "Flowers" cell

// Create a custom note with different properties
const note2 = new Note('Don\'t forget to send these out by Friday!');
note2.setAuthor('Party Planner');
note2.setWidth(200);
note2.setHeight(100);
note2.setBackgroundColor('#FFE4B5'); // Moccasin color
worksheet.insertNote(0, 0, note2); // Add note to "Invitations" cell

// Create a visible note (shows by default instead of on hover)
const note3 = new Note('URGENT: Need to confirm numbers');
note3.setAuthor('Caterer');
note3.setVisible(true);
note3.setBackgroundColor('#FFB6C1'); // Light pink
note3.setWidth(180);
worksheet.insertNote(4, 0, note3); // Add note to "Menu" cell

// Create a note with alt text for accessibility
const note4 = new Note('Check his availability');
note4.setAltText('Note about checking Peter\'s availability');
worksheet.insertNote(5, 0, note4); // Add note to "Peter" cell

// Save the workbook
workbook.save('notes.xlsx');

console.log('✓ Created notes.xlsx');
console.log('  Open the file in Excel to see the notes!');
console.log('  (Hover over cells with red triangles to see notes)');
