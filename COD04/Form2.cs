using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace COD02
{
  public partial class Form2 : Form
  {
    DataTable dt; // 資料表
    string id; // 代碼

    public Form2(DataTable dt, string id)
    {
      InitializeComponent();
      this.dt = dt;
      this.id = id;
    }

    private void Form2_Load(object sender, EventArgs e)
    {
      this.Text = dt.TableName;
      bindingSource1.DataSource = dt;
      textBox1.DataBindings.Add(new Binding("Text", bindingSource1, "代碼", true));
      textBox2.DataBindings.Add(new Binding("Text", bindingSource1, "站點代號", true));
      textBox3.DataBindings.Add(new Binding("Text", bindingSource1, "場站名稱", true));
      textBox4.DataBindings.Add(new Binding("Text", bindingSource1, "經度", true));
      textBox5.DataBindings.Add(new Binding("Text", bindingSource1, "緯度", true));
      textBox6.DataBindings.Add(new Binding("Text", bindingSource1, "地址", true));
    }

    private void button2_Click(object sender, EventArgs e)
    {
      bindingSource1.MovePrevious();
    }

    private void button3_Click(object sender, EventArgs e)
    {
      bindingSource1.MoveNext();
    }

    private void button1_Click(object sender, EventArgs e)
    {
      System.Diagnostics.Process.Start("https://www.google.com/maps/place/" + textBox5.Text + "," + textBox4.Text);
    }
  }
}
