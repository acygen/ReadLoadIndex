using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Security.Cryptography;
using Newtonsoft.Json;
using ExcelDataReader;
using System.Runtime.InteropServices;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ReadLoadIndex
{
    public static class EXCELHelper
    {
        public static List<System.Drawing.Color> stateColors;

        public static AllPlayerData data;
        
        
        private static void AddStateColors()
        {
            stateColors = new List<System.Drawing.Color>();
            stateColors.Add(System.Drawing.Color.FromArgb(178, 178, 178));
            stateColors.Add(System.Drawing.Color.FromArgb(255, 134, 134));
            stateColors.Add(System.Drawing.Color.FromArgb(255, 173, 95));
            stateColors.Add(System.Drawing.Color.FromArgb(151, 168, 255));
            stateColors.Add(System.Drawing.Color.FromArgb(190, 190, 190));
            stateColors.Add(System.Drawing.Color.FromArgb(185, 120, 100));
            stateColors.Add(System.Drawing.Color.FromArgb(135, 135, 135));
            stateColors.Add(System.Drawing.Color.FromArgb(172, 255, 167));
            stateColors.Add(System.Drawing.Color.FromArgb(215, 213, 255));

        }

        public static void SaveExcel(AllPlayerData allPlayerData, string defaultName = "")
        {
            data = allPlayerData;
            bool isSuccess = true;
            string filePath = "";
            OpenFileName ofn = new OpenFileName();

            ofn.structSize = Marshal.SizeOf(ofn);

            //ofn.filter = "All Files\0*.*\0\0";
            //ofn.filter = "Image Files(*.jpg;*.png)\0*.jpg;*.png\0";
            //ofn.filter = "Txt Files(*.txt)\0*.txt\0";

            //ofn.filter = "Word Files(*.docx)\0*.docx\0";
            //ofn.filter = "Word Files(*.doc)\0*.doc\0";
            //ofn.filter = "Word Files(*.doc:*.docx)\0*.doc:*.docx\0";

            //ofn.filter = "Excel Files(*.xls)\0*.xls\0";
            ofn.filter = "Excel Files(*.xlsx)\0*.xlsx\0";  //指定打开格式
                                                           //ofn.filter = "Excel Files(*.xls:*.xlsx)\0*.xls:*.xlsx\0";
                                                           //ofn.filter = "Excel Files(*.xlsx:*.xls)\0*.xlsx:*.xls\0";

            ofn.file = new string(new char[256]);

            ofn.maxFile = ofn.file.Length;

            ofn.fileTitle = new string(new char[64]);

            ofn.maxFileTitle = ofn.fileTitle.Length;

            //ofn.fileTitle = "B1狼狗智克yls200w";

            ofn.initialDir = Application.StartupPath;//默认路径

            ofn.title = "选择保存路径";

            ofn.defExt = "xlsx";
            ofn.file = defaultName;
            //注意 一下项目不一定要全选 但是0x00000008项不要缺少
            ofn.flags = 0x00080000 | 0x00001000 | 0x00000800 | 0x00000200 | 0x00000008;//OFN_EXPLORER|OFN_FILEMUSTEXIST|OFN_PATHMUSTEXIST| OFN_ALLOWMULTISELECT|OFN_NOCHANGEDIR

            isSuccess = DllTest.GetSaveFileName(ofn);
            filePath = ofn.file.Replace("\\", "/");

            //打开windows框
            if (isSuccess)
            {
                //TODO

                //把文件路径格式替换一下
                //ofn.file = ofn.file.Replace("\\", "/");
                //Debug.Log(ofn.file);

                FileInfo newFile = new FileInfo(filePath);
                if (newFile.Exists)
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(filePath);
                }
               CreateExcel(newFile);
                       



            }

        }
        private static void CreateExcel(FileInfo newFile)
        {
            int debugPos = 0;
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    ExcelWorksheet worksheet0 = package.Workbook.Worksheets.Add("BOX");
                    int lineIdx = 2;
                    //for(int i = 0; i < data.allUnitDic.Count; i++)
                    int i = 0;

                    foreach (var pair in data.allUnitDic)
                    {
                        //worksheet0.Cells[1, 1 + i * 7, 1, 1 + (i + 1) * 7-1].Merge = true;
                        worksheet0.Cells[1, 2 + i * 7].Value = $"{pair.Key}({pair.Value})";
                        int unitid = pair.Key;
                        //for(int j = 0; j < data.playerDatas.Count; j++)

                        i++;


                    }
                    i = 0;
                    foreach (var pair2 in data.playerDatas)
                    {
                        worksheet0.Cells[lineIdx, 1].Value = pair2.name;
                        i = 0;

                        foreach (var pair in data.allUnitDic)
                        {
                            //worksheet0.Cells[1, 1 + i * 7, 1, 1 + (i + 1) * 7-1].Merge = true;
                            //worksheet0.Cells[1, 2 + i * 7].Value = $"{pair.Key}({pair.Value})";
                            int unitid = pair.Key;

                            //for(int j = 0; j < data.playerDatas.Count; j++)




                            string[] ppp = pair2.boxDic.TryGetValue(unitid, out var vv) ? vv : new string[7] { "0", "0", "0", "0", "0", "0", "0" };

                            worksheet0.Cells[lineIdx, i * 7 + 2].Value = ppp[0];
                            worksheet0.Cells[lineIdx, i * 7 + 3].Value = ppp[1];
                            worksheet0.Cells[lineIdx, i * 7 + 4].Value = ppp[2];
                            worksheet0.Cells[lineIdx, i * 7 + 5].Value = ppp[3];
                            worksheet0.Cells[lineIdx, i * 7 + 6].Value = ppp[4];
                            worksheet0.Cells[lineIdx, i * 7 + 7].Value = ppp[5];
                            worksheet0.Cells[lineIdx, i * 7 + 8].Value = ppp[6];
                            i++;
                        }

                        lineIdx++;

                    }
                    
                    package.Save();
                }

            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("ERROR:[" + debugPos + "]" + ex.Message + ex.StackTrace);

            }
        }

        

      
        

    }
}
