using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using System.IO;

namespace Candice
{
    public class ImportCSV
    {
        private string importFilePath;
        private string exportFilePath;
        private ArrayList rowAL;
        private Encoding encoding;

        /// <summary>
        /// 文件流读取文件
        /// </summary>
        /// <param name="filePath">服务器上的文件路径</param>
        public ImportCSV(string importFilePath,string exportFilePath)
        {
            this.rowAL = new ArrayList();
            this.importFilePath = importFilePath;
            this.exportFilePath = exportFilePath;
            this.encoding = Encoding.UTF8;
            LoadCSVFile();
        }

        /// <summary>
        /// 载入CSV文件
        /// </summary>
        private void LoadCSVFile()
        {
            if (string.IsNullOrEmpty(this.importFilePath))
            {
                throw new Exception("请指定要载入的CSV文件名");
            }
            else if (!File.Exists(this.importFilePath))
            {
                throw new Exception("指定的CSV文件不存在");
            }
            else
            {
                if (string.IsNullOrEmpty(this.encoding.ToString()))
                {
                    this.encoding = Encoding.UTF8;
                }
            }

            FileStream fs = new FileStream(importFilePath, FileMode.Open);     //参数   文件路径，模式

            StreamReader sr = new StreamReader(fs, this.encoding, true);    //参数  文件流,编码,是否自动检测BOM (Exception:System.IO.FileNotFoundException)
            string csvDataLine = "";

            while (true)
            {
                string fileDataLine;

                fileDataLine = sr.ReadLine();
                if (fileDataLine == null)
                    break;
                if (csvDataLine == "")
                    csvDataLine = fileDataLine;
                else
                {
                    csvDataLine += "\r\n" + fileDataLine;
                }
                //如果包含偶数个引号，说明该行数据中出现回车符或包含逗号
                if (!IfOddQuota(csvDataLine))
                {
                    AddNewDataLine(csvDataLine);
                    csvDataLine = "";
                }

            }

            fs.Dispose();    //关闭流
            sr.Dispose();    //关闭流

            ExportCSVContent();   //输出内容

        }

        /// <summary>
        /// 获取总行数
        /// </summary>
        public int RowCount
        {
            get
            {
                return this.rowAL.Count;
            }
        }

        /// <summary>
        /// 获取总列数
        /// </summary>
        public int ColCount
        {
            get
            {
                int maxCol = 0;
                for (int i = 0; i < this.rowAL.Count; i++)
                {
                    ArrayList colAL = (ArrayList)this.rowAL[i];
                    maxCol = (maxCol > colAL.Count) ? maxCol : colAL.Count;
                }
                return maxCol;
            }
        }

        /// <summary>
        /// 输出CSV内容
        /// </summary>
        public void ExportCSVContent()
        {
            FileStream fs = new FileStream(exportFilePath, FileMode.Create);
            StreamWriter wr = new StreamWriter(fs);

            foreach (var item in rowAL)
            {
                wr.WriteLine(item);
            }

            wr.Flush();
            wr.Dispose();
            fs.Dispose();
        }

        /// <summary>
        /// 判断字符串是否包含奇数个引号
        /// </summary>
        /// <param name="dataLine">数据行</param>
        /// <returns>奇数为true</returns>
        private bool IfOddQuota(string dataLine)
        {
            int quotaCount = 0;
            bool oddQuota = false;

            for (int i = 0; i < dataLine.Length; i++)
            {
                if (dataLine[i] == '\"')
                {
                    quotaCount++;
                }
            }

            if (quotaCount % 2 == 1)
                oddQuota = true;

            return oddQuota;
        }

        /// <summary>
        /// 判断是否以奇数个引号开始
        /// </summary>
        /// <param name="dataCell"></param>
        /// <returns></returns>
        private bool IfOddStartQuota(string dataCell)
        {
            int quotaCount = 0;
            bool oddQuota = false;

            for (int i = 0; i < dataCell.Length; i++)
            {
                if (dataCell[i] == '\"')
                {
                    quotaCount++;
                }
                else
                {
                    break;
                }
            }

            if (quotaCount % 2 == 1)
            {
                oddQuota = true;
            }

            return oddQuota;
        }

        /// <summary>
        /// 判断是否以奇数个引号结尾
        /// </summary>
        /// <param name="dataCell"></param>
        /// <returns></returns>
        private bool IfOddEndQuota(string dataCell)
        {
            int quotaCount = 0;
            bool oddQuota = false;

            for (int i = dataCell.Length - 1; i > 0; i--)
            {
                if (dataCell[i] == '\"')
                    quotaCount++;
                else
                {
                    break;
                }
            }

            if (quotaCount % 2 == 1)
            {
                oddQuota = true;
            }

            return oddQuota;
        }

        /// <summary>
        /// 添加新的数据行
        /// </summary>
        /// <param name="newDataLine"></param>
        private void AddNewDataLine(string newDataLine)
        {
            ArrayList colAL = new ArrayList();
            string dateLine = "";
            string[] dataArray = newDataLine.Split(',');
            bool oddStartQuota = false;   //是否以奇数个引号开始

            string cellData = "";

            for (int i = 0; i < dataArray.Length; i++)
            {
                if (oddStartQuota)
                {
                    //因为前面用逗号分割,所以要加上逗号
                    cellData += "," + dataArray[i];
                    //是否以奇数个引号结尾
                    if (IfOddEndQuota(dataArray[i]))
                    {
                        colAL.Add(GetHandleData(cellData));
                        oddStartQuota = false;
                        continue;
                    }
                    else
                    {
                        oddStartQuota = true;
                        cellData = dataArray[i];
                        continue;
                    }
                }
                else
                {
                    colAL.Add(GetHandleData(dataArray[i]));
                }
            }
            if (oddStartQuota)
            {
                throw new Exception("数据格式有问题");
            }
            for (int i = 0; i < dataArray.Length; i++)
            {
                dateLine += (dataArray[i]+',');
            }
            rowAL.Add(dateLine.Substring(0, dateLine.Length - 1));
        }

        /// <summary>
        /// 去掉格子的首尾引号，把双引号变成单引号
        /// </summary>
        /// <param name="csvCellData"></param>
        /// <returns></returns>
        private string GetHandleData(string csvCellData)
        {
            if (csvCellData == "")
                return "";

            if (IfOddStartQuota(csvCellData))
            {
                if (IfOddEndQuota(csvCellData))
                    return csvCellData.Substring(1, csvCellData.Length - 2).Replace("\"\"", "\"");   //去掉收尾引号,把双引号换成单引号
                else
                {
                    throw new Exception("引号无法匹配" + csvCellData);
                }
            }
            else
            {
                if (csvCellData.Length > 2 && csvCellData[0] == '\"')
                    csvCellData = csvCellData.Substring(1, csvCellData.Length - 2).Replace("\"\"", "\"");
            }

            return csvCellData;
        }
    }
}
