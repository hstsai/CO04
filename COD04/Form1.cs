using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel; // 必須先加入參考Microsoft.office.interop.Excel
using Word = Microsoft.Office.Interop.Word; // 必須先加入參考Microsoft.office.interop.Word

namespace COD02
{
  public partial class Form1 : Form
  {
    DataSet ds; // DataSet儲存匯入資料

    public Form1()
    {
      InitializeComponent();
      ds = new DataSet(); // 建立DataSet物件
    }

    // 選擇檔案
    private void cmdBrowse_Click(object sender, EventArgs e)
    {
      if (openFileDialog1.ShowDialog() == DialogResult.OK) // 選擇Excel資料檔
      {
        txtFilePath.Text = openFileDialog1.FileName;
      }
    }

    // 讀取Excel檔
    private void cmdOpenRead_Click(object sender, EventArgs e)
    {
      string dataSource = txtFilePath.Text; // Excel檔案名

      if (string.IsNullOrEmpty(dataSource))
      {
        MessageBox.Show("未選擇資料來源檔案。");
        return;
      }

      // 使用Office Interop
      COMReadExcel(dataSource);

      MessageBox.Show("讀取Excel成功!");

      // 啟用匯出
      button1.Enabled = true;
    }

    // Office Interop
    private void COMReadExcel(string fileName)
    {
      var app = new Excel.Application(); // 建立Excel應用程式
      //app.Visible = true; // 顯示Excel應用程式
      var wb = app.Workbooks.Open(fileName); // 讀取Excel活頁簿
      ds.Reset(); // 重置DataSet物件
      foreach (Excel.Worksheet ws in wb.Sheets) // 取出工作表
      {
        var dt = new DataTable(ws.Name); // 建立DataTable物件
        foreach (var c in ws.UsedRange.Rows[1].Cells) // 取得Excel欄位標題
        {
          dt.Columns.Add(c.Value);
        }
        for (int r = 2; r <= ws.UsedRange.Rows.Count; r++) // 取得Excel使用區塊
        {
          var nr = dt.NewRow(); // 加入資料錄

          foreach (var c in ws.UsedRange.Rows[r].Cells) // 填入欄位內容
          {
            nr[c.Column - 1] = c.Value;
          }

          dt.Rows.Add(nr); // 加入DataTable
        }

        ds.Tables.Add(dt); // 加入DataTable
        comboBox1.Items.Add(dt.TableName); // 加入ComboBox項目
      }

      app.Quit(); // 結束Excel
    }

    // 選取工作表
    private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
    {
      var tn = comboBox1.SelectedItem.ToString(); // 取出選取工作表
      this.Text = tn; // 設定視窗標題
      var dt = ds.Tables[tn]; // 取出指定DataTable

      bindingSource1.DataSource = dt; // 設定bindingSource1資料來源為dt
      dgvDataList.DataSource = bindingSource1; // 設定dgvDataList資料來源為bindingSource1
    }

    // 匯出Word
    private void button1_Click(object sender, EventArgs e)
    {
      saveFileDialog1.Filter = "Word文件(*.docx)|*.docx";
      if (saveFileDialog1.ShowDialog() == DialogResult.OK)
      {
        Object oMissing = System.Reflection.Missing.Value; // Office Missing
        Object oFalse = false; // Office False
        var wrdApp = new Word.Application(); // 建立Word應用程式 https://msdn.microsoft.com/zh-tw/library/bb157880.aspx
        wrdApp.Visible = true; // 顯示應用程式
        var wrdDoc = wrdApp.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing); // 建立Word文件
        
        // 建立郵件合併資料檔
        Object oName = "Temp.doc"; // 暫存文件資料檔
        Object oHeader = "代碼, 站點代號, 場站名稱, 緯度, 經度, 地址"; // 建立欄位標題
        wrdDoc.MailMerge.CreateDataSource(ref oName, ref oMissing, ref oMissing, ref oHeader, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing); // 建立郵件合併資料來源
        var oDataDoc = wrdApp.Documents.Open(ref oName, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing); // 開啟暫存檔
        
        for (int r = 0; r < ds.Tables[comboBox1.SelectedItem.ToString()].Rows.Count; r++) // 填入Word文件表格
        {
          for (int c = 0; c < ds.Tables[comboBox1.SelectedItem.ToString()].Columns.Count; c++)
          {
            var data = ds.Tables[comboBox1.SelectedItem.ToString()].Rows[r][c].ToString();
            oDataDoc.Tables[1].Cell(r + 2, c + 1).Range.InsertAfter(data);
          }
          if (r < ds.Tables[comboBox1.SelectedItem.ToString()].Rows.Count - 1)
          {
            oDataDoc.Tables[1].Rows.Add(ref oMissing);
          }
        }
        oDataDoc.Save(); // 存檔
        oDataDoc.Close(ref oFalse, ref oMissing, ref oMissing); // 關閉檔案

        // 加入頁首
        var wrdSection = wrdDoc.Sections[1];
        var hr = wrdSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range; // 取得頁首Range
        hr.Font.Name = "新細明體"; // 字型名稱
        hr.Font.Size = 14; // 字型大小
        hr.Text = comboBox1.SelectedItem.ToString(); // 顯示文字
        hr.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; // 置中對齊

        // 加入頁尾
        var fr = wrdSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range; // 取得頁尾Range
        fr.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
        fr.Font.Name = "新細明體"; // 字型名稱
        fr.Font.Size = 10; // 字型大小
        fr.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; // 置中對齊
        fr.Select(); // 選取頁尾
        var wrdSelection = wrdApp.Selection; // 取出頁尾選取區
        wrdSelection.TypeText("第");
        wrdSelection.Fields.Add(wrdSelection.Range, Word.WdFieldType.wdFieldSection); // 目前節數=>頁數
        wrdSelection.TypeText("頁/共");
        wrdSelection.Fields.Add(wrdSelection.Range, Word.WdFieldType.wdFieldNumPages); // 總頁數
        wrdSelection.TypeText("頁");

        // 插入合併資料
        wrdDoc.Select(); // 選取Word文件
        wrdSelection = wrdApp.Selection; // 取出文件選取區或插入點
        wrdSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
        var wrdMailMerge = wrdDoc.MailMerge; // 取出郵件合併
        var wrdMergeFields = wrdMailMerge.Fields; // 取出郵件合併欄位
        wrdSelection.TypeText("代碼：");
        wrdMergeFields.Add(wrdSelection.Range, "代碼");
        wrdSelection.TypeText("　站點代號：");
        wrdMergeFields.Add(wrdSelection.Range, "站點代號");
        wrdSelection.TypeParagraph();

        wrdSelection.TypeText("場站名稱：");
        wrdMergeFields.Add(wrdSelection.Range, "場站名稱");
        wrdSelection.TypeParagraph();

        wrdSelection.TypeText("緯度：");
        wrdMergeFields.Add(wrdSelection.Range, "緯度");
        wrdSelection.TypeText("　經度：");
        wrdMergeFields.Add(wrdSelection.Range, "經度");
        wrdSelection.TypeParagraph();

        wrdSelection.TypeText("地址：");
        wrdMergeFields.Add(wrdSelection.Range, "地址");
        wrdSelection.TypeParagraph();

        wrdMailMerge.Destination = Word.WdMailMergeDestination.wdSendToNewDocument; // 設定合併結果目標
        wrdMailMerge.Execute(ref oFalse); // 執行郵件合併

        wrdDoc.Saved = true; // 指定文件是否存檔
        wrdDoc.Close(ref oFalse, ref oMissing, ref oMissing); // 關閉指定文件

        File.Delete("Temp.doc"); // 刪除暫存資料檔
      }
    }

    // 顯示單筆資料
    private void dgvDataList_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
    {
      var dt = ds.Tables[comboBox1.SelectedItem.ToString()]; // 取得資料表
      var id = dgvDataList[0, e.RowIndex].Value.ToString(); // 取得目前選取代號
      var f2 = new Form2(dt, id); // 傳入Form2
      f2.ShowDialog(); // 顯示Form2
    }
  }
}