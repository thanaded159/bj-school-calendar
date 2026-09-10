const SOURCE_CONFIG = Object.freeze([
  { id: '1i5SnD7uRGjwUioVFC4nkiBYvxxjEhRBssg8jnoYEZls', division: 'ฝ่ายบริหารงานทั่วไป', divisionKey: 'general' },
  { id: '1_zCxLke8OXnXdBbuLUz1aIfsvvw1zsk7hXbpDvHWebk', division: 'ฝ่ายบริหารงานบุคคลและกิจการนักเรียน', divisionKey: 'personnel' },
  { id: '1On8aCDfQlu3T5QfNPzcNVOZTNyYSpSatn2QO_oEIhDs', division: 'ฝ่ายบริหารงานวิชาการ', divisionKey: 'academic' },
  { id: '1OR_S9HfA_gcw3fBCyC2o4-RSvwaMLdWd79pd9cls630', division: 'ฝ่ายบริหารงานงบประมาณ', divisionKey: 'budget' }
]);

const THAI_MONTHS = Object.freeze({'ม.ค.':1,'ก.พ.':2,'มี.ค.':3,'เม.ย.':4,'พ.ค.':5,'มิ.ย.':6,'ก.ค.':7,'ส.ค.':8,'ก.ย.':9,'ต.ค.':10,'พ.ย.':11,'ธ.ค.':12});

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ปฏิทินปฏิบัติงานและกิจกรรมโรงเรียน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getCalendarData() {
  const events = [], seen = new Set(), sources = [];
  SOURCE_CONFIG.forEach(source => {
    try {
      const spreadsheet = SpreadsheetApp.openById(source.id);
      let count = 0;
      spreadsheet.getSheets().forEach(sheet => {
        const values = sheet.getDataRange().getValues();
        if (values.length < 4) return;
        const title = String(values[0][0] || '');
        const academicYear = getAcademicYear_(title, sheet.getName());
        const semester = getSemester_(title, sheet.getName());
        if (!academicYear) return;
        for (let rowIndex = 3; rowIndex < values.length; rowIndex++) {
          const row = values[rowIndex];
          for (let column = 0; column < row.length; column += 3) {
            const rawDate = row[column], rawActivity = row[column + 1], rawOwner = row[column + 2];
            if (isBlank_(rawDate) || isBlank_(rawActivity)) continue;
            const activity = cleanText_(rawActivity), owner = cleanText_(rawOwner) || 'ไม่ระบุ';
            const date = normalizeDate_(rawDate, academicYear, semester);
            const key = [source.divisionKey, academicYear, semester, date.dateText, activity, owner].join('|');
            if (seen.has(key)) continue;
            seen.add(key);
            events.push({id:'',academicYear,semester,dateText:date.dateText,start:date.start,activity,owner,division:source.division,divisionKey:source.divisionKey,sourceSheet:sheet.getName(),sourceUrl:spreadsheet.getUrl()});
            count++;
          }
        }
      });
      sources.push({divisionKey:source.divisionKey,ok:true,count});
    } catch (error) {
      sources.push({divisionKey:source.divisionKey,ok:false,count:0,message:String(error.message || error)});
    }
  });
  events.sort((a,b)=>String(b.start||'').localeCompare(String(a.start||'')));
  events.forEach((event,index)=>event.id=`e${index+1}`);
  return {events,updatedAt:new Date().toISOString(),sources,partial:sources.some(source=>!source.ok)};
}

function getAcademicYear_(title, sheetName) {
  const match = `${title} ${sheetName}`.match(/(25\d{2})/);
  return match ? Number(match[1]) : null;
}

function getSemester_(title, sheetName) {
  const match = title.match(/ภาคเรียนที่\s*([12])/);
  return match ? Number(match[1]) : (String(sheetName).startsWith('2') ? 2 : 1);
}

function normalizeDate_(value, academicYear, semester) {
  if (value instanceof Date && !isNaN(value)) {
    const sourceYear=value.getFullYear();
    const buddhistYear=sourceYear>=2500?sourceYear:(sourceYear>=1900&&sourceYear%100===academicYear%100?academicYear:sourceYear+543);
    const month=value.getMonth()+1, day=value.getDate();
    return {dateText:`${day} ${Object.keys(THAI_MONTHS)[month-1]} ${String(buddhistYear).slice(-2)}`,start:toIsoDate_(buddhistYear-543,month,day)};
  }
  const dateText=cleanText_(value).replace(/–/g,'-');
  const match=dateText.match(/(\d{1,2})(?:\s*-\s*\d{1,2})?\s*(ม\.ค\.|ก\.พ\.|มี\.ค\.|เม\.ย\.|พ\.ค\.|มิ\.ย\.|ก\.ค\.|ส\.ค\.|ก\.ย\.|ต\.ค\.|พ\.ย\.|ธ\.ค\.)\s*(\d{2,4})?/);
  if(!match) return {dateText,start:null};
  const day=Number(match[1]),month=THAI_MONTHS[match[2]];
  const buddhistYear=match[3]?(Number(match[3])<100?Number(match[3])+2500:Number(match[3])):academicYear+(semester===2&&month<=4?1:0);
  return {dateText,start:toIsoDate_(buddhistYear-543,month,day)};
}

function toIsoDate_(year,month,day) {
  const date=new Date(year,month-1,day);
  if(date.getFullYear()!==year||date.getMonth()!==month-1||date.getDate()!==day) return null;
  return Utilities.formatDate(date,'Asia/Bangkok','yyyy-MM-dd');
}

function cleanText_(value){return String(value==null?'':value).replace(/\s+/g,' ').trim();}
function isBlank_(value){return value===null||value===undefined||value==='';}
