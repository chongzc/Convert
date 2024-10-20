const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const handlebars = require('handlebars');
const htmlDocx = require('html-docx-js');
//const { PDFDocument } = require('pdf-lib');

const jsonData = {
    "Definitions-and-Interpretation": "1",
    "date": "3 OCTORBER, 2024",
    "company": "Amiphotoshop Sdn Bhd",
    "companyNumber": "12345678",
    "address": "123, Jalan XYZ, 50000 Kuala Lumpur, Malaysia",
    "telephone": "0123456789",
    "email": "XJ@gmail.com",
    "attn": "Composition of board",
    "shareholders": [
        {
            "name": "Abu",
            "icPassportCompanyNumber": "IC123456789",
            "address": "No. 1, Jalan ABC, 70000 Seremban, Malaysia",
            "telephone": "01928282828",
            "email": "hohoh@asdf.com",
            "attn": "asdfasdf",
            "contribution": "donation",
            "noOfShares": 500000,
            "percentage": "50.0"
        },
        {
            "name": "Ali",
            "icPassportCompanyNumber": "P12345678",
            "address": "No. 2, Jalan DEF, 80000 Johor Bahru, Malaysia",
            "telephone": "01928282828",
            "email": "hohoh@asdf.com",
            "attn": "asdfasdf",
            "contribution": "donation",
            "noOfShares": 500000,
            "percentage": "50.0"
        }
    ],
    "increaseInShareCapital": "Yes",
    "businessOfCompany": "Marketing services and consulting",
    "totalDirectors": "4",
    "directorByShareholder": "Yes",
    "chairmanAppoint": "Yes",
    "chairman": "Ali",
    "chairmanCastingVote": "Yes",
    "quorumForBoardMeeting": "50",
    "reservedMatters": [
        {
            "matter": "Approval of any substantial change in the nature of the business",
        },
        {
            "matter": "Issuance of new shares",
        },
        {
            "matter":"Appointment or removal of directors"
        }
    ],
    "quorumForShareholdersMeeting": "50",
    "lockInPeriod": "5",
    "restrictionOnTransfer": "Yes",
    "restrictionDetails": "All shares must be offered to the existing shareholders before they can be transferred to any third party.",
    "includedTagAlongRight": "Yes",
    "includedDragAlongRight": "Yes",
    "deadlockStatus": "Agree",
    "nonCompeteStatus": "Agree",
    "rolesAndResponsibilities": [
        {
            "name": "Party A",
            "role": "Marketing and Sales",
            "responsibilities": [
                "Develop and implement marketing strategies.",
                "Manage the sales team and oversee sales operations.",
                "Coordinate promotional campaigns and public relations activities.",
                "Monitor market trends and competitor activities."
            ]
        },
        {
            "name": "Party B",
            "role": "Finance and Operations",
            "responsibilities": [
                "Oversee financial planning and budgeting.",
                "Manage the company’s financial resources.",
                "Ensure legal and regulatory compliance.",
                "Supervise the day-to-day operations of the company."
            ]
        },
        {
            "name": "Party C",
            "role": "Technical Development",
            "responsibilities": [
                "Lead the research and development team.",
                "Oversee product design and technical development.",
                "Implement new technologies to improve company products.",
                "Ensure the security and integrity of IT systems."
            ]
        }
    ],
    "transferor-during-lock-in": "1000000",
    "transferor-after-lock-in": "1200000",
    "deadlock-during-lock-in": "1000000",
    "deadlock-after-lock-in": "1200000",
    "default-during-lock-in": "900000",
    "default-after-lock-in": "1100000",
    "beneficiary": "Abu",
    "consideration": "500000",
    "auto-transfer-proportion": "50"
};



//Calculate totals
let totalValue = 0;
let totalPercentage = 0;

jsonData.shareholders.forEach(shareholder => {
    totalValue += shareholder.noOfShares;
    totalPercentage += parseFloat(shareholder.percentage);
});

jsonData.totalValue = totalValue;
jsonData.totalPercentage = totalPercentage.toFixed(1);


const A4_PAGE_HEIGHT = 908; 
const INCH_TO_PX = 99; 
const MARGIN_TOP = INCH_TO_PX; 
const MARGIN_BOTTOM = INCH_TO_PX; 
const USABLE_HEIGHT = A4_PAGE_HEIGHT - (MARGIN_TOP + MARGIN_BOTTOM);

const idsToCheck = [
    "_Toc153272493",
    "_Toc153272494",
    "_Toc153272495",
    "_Toc153272496",
    "_Toc153272497",
    "_Toc153272498",
    "_Toc153272499",
    "_Toc153272500",
    "_Toc153272501",
    "_Toc153272502",
    "_Toc153272503",
    "_Toc153272504",
    "_Toc153272505",
    "_Toc153272506",
    "_Toc153272507",
    "_Toc153272508",
    "_Toc153272509",
    "_Toc153272510",
    "_Toc153272512",
    "_Toc153272513",
    "_Toc153272514",
    "_Toc153272515",
    "_Toc153272516",
    "_Toc153272517",
    "_Toc153272518",
    "_Toc153272519",
    "_Toc153272520",
    "_Toc153272521",
    "_Toc153272522",
    "_Toc153272523",
    "_Toc153272524"
];

(async () => {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();

    const htmlContent = fs.readFileSync('./SHA.html', 'utf8');
    await page.setContent(htmlContent);

    await page.pdf({
        format: 'A4',
        margin: {
            top: '1in',
            right: '1in',
            bottom: '1in',
            left: '1in'
        }
    });

    for (const id of idsToCheck) {
        const element = await page.$(`#${id}`);
        if (element) {
            const boundingBox = await element.boundingBox(); 
        
            if (boundingBox) {
                const topOffset = boundingBox.y; 
                const adjustedTopOffset = topOffset - MARGIN_TOP;
                const pageNumber = Math.floor(adjustedTopOffset / USABLE_HEIGHT) + 1;
    
                jsonData[id] = pageNumber;
            }
        } else {
            console.log(`Not Found: ${id}`);
        }
    }

    await browser.close();

    let ShaHtml = fs.readFileSync('./SHA.html', 'utf-8');
    const ShaTemplate = handlebars.compile(ShaHtml);
    const filledShaHtml = ShaTemplate(jsonData);

    await generateShaPDF(filledShaHtml);
})();


//read SHA checklist template
let checklistHtml = fs.readFileSync('./SHA Checklist.html', 'utf8');
const checkListTemplate = handlebars.compile(checklistHtml);
const filledCheckListHtml = checkListTemplate(jsonData);

//Read SHA template
let ShaHtml = fs.readFileSync('./SHA.html', 'utf-8');
const ShaTemplate = handlebars.compile(ShaHtml);
const filledShaHtml = ShaTemplate(jsonData);

// Generate file name
function getFileName(companyName, extension) {
    const now = new Date();
    const date = now.toISOString().split('T')[0];
    const time = now.toTimeString().split(' ')[0].replace(/:/g, '-');
    return `${companyName}-${date}_${time}.${extension}`;
}

// Create Check List PDF
async function createCheckListPDF() {
    try {
        const browser = await puppeteer.launch();
        const page = await browser.newPage();
        await page.setContent(filledCheckListHtml, { waitUntil: 'networkidle0' });

        const outputFileName = getFileName(jsonData.company + ' check list ', 'pdf');
        const outputPath = path.join(__dirname, outputFileName);

        await page.pdf({
            path: outputPath,
            format: 'A4',
            landscape: true,
            printBackground: true,
            margin: {
                top: '10mm',
                right: '10mm',
                bottom: '10mm',
                left: '10mm'
            },
            displayHeaderFooter: true,
            footerTemplate: `
                <div style="width: 100%; text-align: right; font-size: 10px; padding-right: 10mm;">
                    <span class="pageNumber"></span>/<span class="totalPages"></span>
                </div>
            `,
            headerTemplate: `<div></div>`,
            scale: 1.2
        });

        console.log(`Check List PDF generated: ${outputFileName}`);
        await browser.close();
    } catch (e) {
        console.error(e);
    }
}

// Create check list docx
async function createCheckListDocx() {
    try {
        const docxBlob = htmlDocx.asBlob(filledCheckListHtml,{orientation: 'landscape'});
        const docxBuffer = Buffer.from(await docxBlob.arrayBuffer());
        const outputFileName = getFileName(jsonData.company + ' check list ', 'docx');
        const outputPath = path.join(__dirname, outputFileName);

        fs.writeFileSync(outputPath, docxBuffer);
        console.log(`Check list DOCX generated: ${outputFileName}`);
    } catch (e) {
        console.error(e);
    }
}

// Generate SHA PDF
async function generateShaPDF(filledShaHtml) {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    await page.setContent(filledShaHtml, { waitUntil: 'networkidle0' });
    const outputFileName = getFileName(jsonData.company + ' Sha', 'pdf');
    const outputPath = path.join(__dirname, outputFileName);

    await page.pdf({
        path: outputPath,
        format: 'A4',
        margin: {
            top: '1in',
            right: '1in',
            bottom: '1in',
            left: '1in'
        },
        displayHeaderFooter: true, 
        headerTemplate: '<div></div>',
        footerTemplate: `
            <div style="font-size:10px; color: black; text-align: center; width: 100%; padding: 0.5in;">
                <span class="pageNumber"></span>
            </div>`,
        printBackground: true,
    });

    console.log(`SHA PDF generated: ${outputPath}`);

    await browser.close();
}

// Create SHA docx
async function createShaDocx() {
    try {
        const docxBlob = htmlDocx.asBlob(filledShaHtml, { orientation: 'portrait' });
        
        const docxBuffer = Buffer.from(await docxBlob.arrayBuffer());
        
        const outputFileName = getFileName(jsonData.company + ' SHA', 'docx');
        const outputPath = path.join(__dirname, outputFileName);
        
        fs.writeFileSync(outputPath, docxBuffer);
        
        console.log(`SHA DOCX generated: ${outputFileName}`);
    } catch (e) {
        console.error(e);
    }
}


async function generateFiles() {
    await createCheckListPDF();
    await createCheckListDocx(); 
    await createShaDocx();
}

generateFiles();