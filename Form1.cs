using System;
using System.IO;
using System.Reflection.PortableExecutable;
using OfficeOpenXml;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace pdfsplit
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                openFileDialog.Title = "Select an Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // 선택한 파일의 경로를 텍스트박스에 표시
                    textBox2.Text = openFileDialog.FileName;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // 엑셀 파일 경로 지정
            string filePath = textBox2.Text;
            textBox1.Clear();
            excelHandling(filePath);

        }

        public void excelHandling(string filePath)
        {
            // 읽고자 하는 열 지정 (예: 3번째 열이면 n = 3)
            int n = 3;

            // EPPlus 라이브러리를 사용하여 엑셀 파일을 읽기
            FileInfo fileInfo = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                // 첫 번째 시트 가져오기
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                int startRow = 6; // 시작 행 (A열의 6번째 행부터)

                // A열의 6번째 행부터 문자열이 있는 행까지 순회
                string tmpPdf = "";
                int startPageNo = 0;
                int finishPageNo = 0;
                int totPage = 0;
                for (int row = startRow; row <= worksheet.Dimension.End.Row; row++)
                {
                    // A열의 현재 행의 값 가져오기
                    string pdfName = worksheet.Cells[row, 1].Text;
                    string outPdfName = worksheet.Cells[row, 2].Text;

                    if (pdfName != tmpPdf)
                    {
                        startPageNo = 0;
                        finishPageNo = 0;
                    }

                    string pageCnt = worksheet.Cells[row, 14].Text;
                    finishPageNo = startPageNo + int.Parse(pageCnt) - 1;


                    // 값 출력 (또는 다른 작업 수행)
                    //Console.WriteLine($"Row {row}: A={cellValue}, N={nColumnValue}");
                    textBox1.AppendText("엑셀파일: " + Path.GetFileName(filePath) + 
                           $"처리행: {row}: PDF명={pdfName}, 페이지수:{pageCnt}, 시작페이지:{startPageNo}, 끝페이지:{finishPageNo}, 출력파일: {outPdfName}.pdf" + Environment.NewLine);
                    int ret = pdfsplit(pdfName, outPdfName, startPageNo, finishPageNo, row);
                    if (ret != 0)
                    {
                        textBox1.AppendText("###### PDF 오류: " + Path.GetFileName(filePath) + 
                            $"처리행: {row}: PDF명={pdfName}, 페이지수:{pageCnt}, 시작페이지:{startPageNo}, 끝페이지:{finishPageNo}, 출력파일: {outPdfName}.pdf" + Environment.NewLine);
                    }


                    startPageNo = finishPageNo + 1;
                    tmpPdf = pdfName;
                }
            }
        }

        public string getPathStr (string fileName)
        {
            return fileName.Replace(Path.GetFileName(fileName), "");
        }

        public int pdfsplit(string pdfName, string outPdfName, int startPageNo, int finishPageNo, int row)
        {
            string filePath = getPathStr ( textBox2.Text) ;
            Console.WriteLine("[ " + filePath + "][" + $"Row {row}: [[ PDF Name={pdfName} ]],  start={startPageNo} finish={finishPageNo}");
            /*
            1. inputpdf 파일명을 만든다.
            2. outpdf 파일명을 만든다.
            3. inputpdf에서 startPageNo부터 finishPageNo까지 페이지를 outPDF파일에 추가 한다.
            4. outputpdf를 저장한다.
            */
            try
            {
                PdfDocument outputPDF = new PdfDocument();
                PdfDocument inputPDF = PdfReader.Open(filePath + pdfName + ".pdf", PdfDocumentOpenMode.Import);
                for(int i= startPageNo; i<= finishPageNo; i++)
                {
                    PdfPage selPage = inputPDF.Pages[i];
                    outputPDF.AddPage(selPage);
                }
                inputPDF.Close();
                outputPDF.Save(filePath + outPdfName + ".pdf");
            }
            catch (Exception e)
            {
                textBox1.AppendText(e.ToString());
                Console.WriteLine(e.ToString());

                return 1;
            }
            return 0;
        }

    }
}
