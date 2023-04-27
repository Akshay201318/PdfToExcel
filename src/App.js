import { useState } from 'react';
import { saveAs } from 'file-saver';
import * as pdfjsLib from 'pdfjs-dist';
import * as XLSX from 'xlsx';
import './Loader.css';

pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.js`; // Set workerSrc to correct URL

const deliveryPartners = {
  BLUEDART: true,
  'Ecom Express': true,
};

function App() {
  const [pdfFile, setPdfFile] = useState(null);
  const [loader, setLoader] = useState(false);

  // Handler for PDF file upload
  const handlePdfFileUpload = (e) => {
    const file = e.target.files;
    setPdfFile(file);
  };
  // Handler for extracting data from PDF and creating Excel
  const handleExtractDataAndCreateExcel = async () => {
    setLoader(true);
    if (pdfFile) {
      let extractedData = [];
      const reader = new FileReader();
      let fileIndex = 0;
      let onloadFileIndex = 0;
      reader.onload = async (e) => {
        const buffer = new Uint8Array(e.target.result);
        const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
        const numPages = pdf.numPages;

        for (let i = 1; i <= numPages; i++) {
          const data = {};
          const page = await pdf.getPage(i);
          const textContent = await page.getTextContent();
          const items = textContent.items;
          data['Name'] = items[2]?.str?.split('-')[0];
          data['Mobile'] = items[4]?.str?.split(',')[0];
          let address = '';
          let index = 6;
          while (true) {
            if (deliveryPartners[items[index].str]) {
              break;
            }
            if (!items[index].str.length) {
              index++;
              continue;
            }
            address += ' ' + items[index].str;
            index++;
          }
          data['Address'] = address;
          address = address.split(',');
          address = address.slice(address.length - 2);
          data['State'] = address[0];
          data['PIN code'] = address[1].split(' ').join('');
          extractedData.push(data);
        }
        if (onloadFileIndex === pdfFile.length - 1) {
          writeXLfile(extractedData);
        } else {
          ++onloadFileIndex;
        }
      };
      reader.readAsArrayBuffer(pdfFile[fileIndex]);
      reader.onloadend = () => {
        if (fileIndex === pdfFile?.length - 1) {
          // writeXLfile();
          return;
        }
        reader.readAsArrayBuffer(pdfFile[++fileIndex]);
      };
    }
  };

  const writeXLfile = (extractedData) => {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(extractedData);

    let colFormatting = [];
    Object.keys(extractedData[0]).forEach(() =>
      colFormatting.push({ wch: 30 })
    );
    worksheet['!cols'] = colFormatting;

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const excelBuffer = XLSX.write(workbook, {
      type: 'array',
      bookType: 'xlsx',
    });
    saveAs(
      new Blob([excelBuffer], { type: 'application/octet-stream' }),
      'extracted_data.xlsx'
    );
    setLoader(false);
    setPdfFile(null);
  };
  return (
    <div
      style={{
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        width: '100vw',
        height: '90vh',
        flexDirection: 'column',
      }}
    >
      <h1
        style={{
          color: '#383e45',
          fontWeight: 500,
        }}
      >
        Convert PDF to EXCEL
      </h1>
      {!pdfFile ? (
        <div
          style={{
            position: 'relative',
            width: 200,
            height: 58,
          }}
        >
          <input
            type='file'
            accept='.pdf'
            onChange={handlePdfFileUpload}
            style={{
              position: 'absolute',
              height: '100%',
              opacity: 0,
              zIndex: 1,
            }}
            multiple={true}
          />
          <div
            style={{
              position: 'absolute',
              padding: 20,
              background: '#e5322d',
              borderRadius: 8,
              color: '#FFF',
              cursor: 'pointer',
            }}
          >
            Select Pdf To Upload
          </div>
        </div>
      ) : (
        <div
          onClick={handleExtractDataAndCreateExcel}
          style={{
            padding: 20,
            background: '#e5322d',
            borderRadius: 8,
            color: '#FFF',
            cursor: 'pointer',
            position: 'relative',
            overflow: 'hidden',
          }}
        >
          {loader && (
            <div class='container'>
              <div class='loader' />
            </div>
          )}
          Download Excel
        </div>
      )}
    </div>
  );
}

export default App;
