/**
 * Sparklines example - demonstrates inline mini-charts in cells
 */

const { Workbook, Format, Sparkline } = require('@nazarkalytiuk/rusc-xlsx');

const workbook = new Workbook();
const worksheet = workbook.addWorksheet('Sparklines');

const headerFmt = new Format().setBold().setBackgroundColor('#4472C4').setFontColor('white');

worksheet.setColumnWidth(0, 15);
worksheet.setColumnWidth(7, 18);

// Write headers
worksheet.write(0, 0, 'Metric', headerFmt);
['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].forEach((day, i) => {
  worksheet.write(0, i + 1, day, headerFmt);
});
worksheet.write(0, 7, 'Trend', headerFmt);

// Row 1: Sales data with line sparkline
worksheet.write(1, 0, 'Sales');
[120, 150, 130, 170, 160, 190].forEach((v, i) => worksheet.write(1, i + 1, v));

const lineSpark = new Sparkline()
  .setRange("Sparklines!B2:G2")
  .setType('line')
  .setSparklineColor('#4472C4')
  .showHighPoint(true)
  .setHighPointColor('#FF0000')
  .showLowPoint(true)
  .setLowPointColor('#00B050');

worksheet.addSparkline(1, 7, lineSpark);

// Row 2: Visitors data with column sparkline
worksheet.write(2, 0, 'Visitors');
[300, 280, 350, 310, 400, 450].forEach((v, i) => worksheet.write(2, i + 1, v));

const colSpark = new Sparkline()
  .setRange("Sparklines!B3:G3")
  .setType('column')
  .setSparklineColor('#ED7D31')
  .showFirstPoint(true)
  .setFirstPointColor('#4472C4')
  .showLastPoint(true)
  .setLastPointColor('#70AD47');

worksheet.addSparkline(2, 7, colSpark);

// Row 3: Profit/loss with win/lose sparkline
worksheet.write(3, 0, 'Profit/Loss');
[10, -5, 15, -3, 20, 8].forEach((v, i) => worksheet.write(3, i + 1, v));

const winLoseSpark = new Sparkline()
  .setRange("Sparklines!B4:G4")
  .setType('winLose')
  .setSparklineColor('#70AD47')
  .showNegativePoints(true)
  .setNegativePointsColor('#FF0000');

worksheet.addSparkline(3, 7, winLoseSpark);

// Row 5+: Sparkline group (multiple sparklines sharing settings)
worksheet.write(5, 0, 'Team', headerFmt);
['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'].forEach((m, i) => {
  worksheet.write(5, i + 1, m, headerFmt);
});
worksheet.write(5, 7, 'Trend', headerFmt);

const teams = ['Alpha', 'Beta', 'Gamma', 'Delta'];
const teamData = [
  [50, 60, 55, 70, 80, 85],
  [40, 45, 50, 48, 55, 60],
  [30, 35, 25, 40, 45, 50],
  [70, 65, 75, 80, 72, 90],
];

teams.forEach((team, i) => {
  worksheet.write(6 + i, 0, team);
  teamData[i].forEach((v, j) => worksheet.write(6 + i, j + 1, v));
});

const groupSpark = new Sparkline()
  .setRange("Sparklines!B7:G10")
  .setType('line')
  .setSparklineColor('#5B9BD5')
  .showMarkers(true)
  .setMarkersColor('#ED7D31')
  .setStyle(3)
  .setGroupMax(true)
  .setGroupMin(true);

worksheet.addSparklineGroup(6, 7, 9, 7, groupSpark);

workbook.save('sparklines.xlsx');
console.log('Created sparklines.xlsx with line, column, win/lose sparklines and sparkline groups');
