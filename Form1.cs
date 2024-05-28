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
                    // ������ ������ ��θ� �ؽ�Ʈ�ڽ��� ǥ��
                    textBox2.Text = openFileDialog.FileName;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // ���� ���� ��� ����
            string filePath = textBox2.Text;
            textBox1.Clear();
            excelHandling(filePath);

        }

        public void excelHandling(string filePath)
        {
            // �а��� �ϴ� �� ���� (��: 3��° ���̸� n = 3)
            int n = 3;

            // EPPlus ���̺귯���� ����Ͽ� ���� ������ �б�
            FileInfo fileInfo = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                // ù ��° ��Ʈ ��������
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                int startRow = 6; // ���� �� (A���� 6��° �����)

                // A���� 6��° ����� ���ڿ��� �ִ� ����� ��ȸ
                string tmpPdf = "";
                int startPageNo = 0;
                int finishPageNo = 0;
                int totPage = 0;
                for (int row = startRow; row <= worksheet.Dimension.End.Row; row++)
                {
                    // A���� ���� ���� �� ��������
                    string pdfName = worksheet.Cells[row, 1].Text;
                    string outPdfName = worksheet.Cells[row, 2].Text;

                    if (pdfName != tmpPdf)
                    {
                        startPageNo = 0;
                        finishPageNo = 0;
                    }

                    string pageCnt = worksheet.Cells[row, 14].Text;
                    finishPageNo = startPageNo + int.Parse(pageCnt) - 1;


                    // �� ��� (�Ǵ� �ٸ� �۾� ����)
                    //Console.WriteLine($"Row {row}: A={cellValue}, N={nColumnValue}");
                    textBox1.AppendText("��������: " + Path.GetFileName(filePath) + 
                           $"ó����: {row}: PDF��={pdfName}, ��������:{pageCnt}, ����������:{startPageNo}, ��������:{finishPageNo}, �������: {outPdfName}.pdf" + Environment.NewLine);
                    int ret = pdfsplit(pdfName, outPdfName, startPageNo, finishPageNo, row);
                    if (ret != 0)
                    {
                        textBox1.AppendText("###### PDF ����: " + Path.GetFileName(filePath) + 
                            $"ó����: {row}: PDF��={pdfName}, ��������:{pageCnt}, ����������:{startPageNo}, ��������:{finishPageNo}, �������: {outPdfName}.pdf" + Environment.NewLine);
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
            1. inputpdf ���ϸ��� �����.
            2. outpdf ���ϸ��� �����.
            3. inputpdf���� startPageNo���� finishPageNo���� �������� outPDF���Ͽ� �߰� �Ѵ�.
            4. outputpdf�� �����Ѵ�.
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
