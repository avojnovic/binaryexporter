using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;

namespace BinaryExporter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            progressBar1.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

                int cant = 0;
                string RTV = "";
                string RTV_Ant = "";
                bool newFile = true;
                

                lblMensaje.Text = "";
                button1.Enabled = false;
                textBox2.Enabled = false;
                progressBar1.Visible = true;
                progressBar1.Minimum = 1;
                StringBuilder sb = new StringBuilder();
                int binaryCol = 2;
                SqlConnection cn = new SqlConnection("server=Marg04;integrated security=yes;database=MonsantoED");

                SqlCommand cmdCount = new SqlCommand(@"
                    SELECT Count(1)
                  FROM [MonsantoED].[dbo].[Informe_Productor] i
                  inner join [MonsantoED].dbo.Support_Incident_Attachment a on a.Informe_Productor_Id=i.Informe_Productor_Id
                  inner join [MonsantoED].dbo.Distribuidores d on  i.centro_de_servicio_id=d.Distribuidores_Id
                  inner join [MonsantoED].dbo.territory t on d.territorio_canal_id=t.territory_id
                  inner join [MonsantoED].dbo.employee e on e.Employee_Id=t.Responsable_Id
                  where i.metricasPOS_Seleccionado=1 ", cn);


                cn.Open();
                SqlDataReader dr = cmdCount.ExecuteReader();

                dr.Read();
                int docs = dr.GetInt32(0);
                progressBar1.Maximum = docs;
                lblMensaje.Text = docs.ToString() + " To Export";
                progressBar1.Value = 1;
                progressBar1.Step = 1;

                dr.Close();
                cn.Close();

                SqlCommand cmd = new SqlCommand(@"
               SELECT Nombre_Distribuidor,
                i.rn_descriptor,Attachment_Binary,RN_Attachment_Binary_FN,RN_Attachment_Binary_FZ
                ,e.Rn_Descriptor,cuit,i.NombreArchivo
                  FROM [MonsantoED].[dbo].[Informe_Productor] i
                  inner join [MonsantoED].dbo.Support_Incident_Attachment a on a.Informe_Productor_Id=i.Informe_Productor_Id
                  inner join [MonsantoED].dbo.Distribuidores d on  i.centro_de_servicio_id=d.Distribuidores_Id
                  inner join [MonsantoED].dbo.territory t on d.territorio_canal_id=t.territory_id
                  inner join [MonsantoED].dbo.employee e on e.Employee_Id=t.Responsable_Id
                  where i.metricasPOS_Seleccionado=1 
                    order by e.Rn_Descriptor,Nombre_Distribuidor", cn);



                cn.Open();

                dr = cmd.ExecuteReader();



                
                while (dr.Read())
                {
                    cant++;
                    lblMensaje.Text = cant.ToString() + " of " + docs.ToString() + " To Export";

                    Byte[] b = new Byte[(dr.GetBytes(binaryCol, 0, null, 0, int.MaxValue))];
                    dr.GetBytes(binaryCol, 0, b, 0, b.Length);
                    System.IO.FileStream fs;

                    string path = textBox2.Text + "/" + dr.GetString(5) + "/" + dr.GetString(0);
                    string filename = dr.GetString(7) + "-" + dr.GetString(3);
                    RTV = dr.GetString(5);
                    if (RTV_Ant == "")
                        RTV_Ant = RTV;


                    if (RTV != RTV_Ant)
                    {
                        WriteCsv(sb, dr, filename, textBox2.Text + "/" + RTV_Ant, newFile, true, false);
                        newFile = true;
                        sb = new StringBuilder();
                    }

                    if (!System.IO.Directory.Exists(path))
                        System.IO.Directory.CreateDirectory(path);


                    string[] files = System.IO.Directory.GetFiles(path);
                    if (files.Count() < 10 || files.Contains(filename.Trim()))
                    {
                        WriteCsv(sb, dr, filename, textBox2.Text + "/" + dr.GetString(5), newFile, false, true);
                        newFile = false;

                        fs = new System.IO.FileStream(path + "/" + filename, System.IO.FileMode.Create, System.IO.FileAccess.Write);
                        fs.Write(b, 0, b.Length);
                        fs.Close();
                    }

   

                    RTV_Ant = dr.GetString(5);

                    if (docs == cant)
                    {
                        WriteCsv(sb, dr, filename, textBox2.Text + "/" + RTV_Ant, newFile, true, false);
                    }


                    progressBar1.PerformStep();
                    Application.DoEvents();
                }



                dr.Close();
                cn.Close();


                lblMensaje.Text = "Export successfully";

            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                progressBar1.Visible = false;
                button1.Enabled = true;
                textBox2.Enabled = true;
            }
        }

        private static void WriteCsv(StringBuilder sb, SqlDataReader dr, string filename, string path, bool newFile, bool write,bool addRow)
        {
            if (newFile)
                sb.AppendLine("RTV;CUIT;POS;Informe;Puntaje");

            if (addRow)
                sb.AppendLine(dr.GetString(5) + ";" + dr.GetString(6) + ";" + dr.GetString(0) + ";" + filename + ";");

            if (write)
            {
                StreamWriter outfile = new StreamWriter(path + @"\PlanillaPuntajes.csv");
                outfile.Write(sb.ToString());
                outfile.Close();
              
            }

        }
    }
}
