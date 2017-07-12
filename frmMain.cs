using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace CargaMasivaTablas
{
	/// <summary>
	/// Descripción breve de Form1.
	/// </summary>
	public class frmMain : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblQuery;
		private System.Windows.Forms.TextBox txtQuery;
		private System.Windows.Forms.Button btnEjecutar;
		private System.Windows.Forms.TextBox txtConectionString;
		private System.Windows.Forms.Label lblConectionString;
		private System.Windows.Forms.Label lblResultado;
		private System.Windows.Forms.TextBox txtResult;
		private System.Windows.Forms.CheckBox chkCalculaNombreTabla;
		private System.Windows.Forms.TextBox txtNombreTabla;
		private System.Windows.Forms.Label lblNombreTabla;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmMain()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		/// <summary>
		/// Limpiar los recursos que se estén utilizando.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.lblQuery = new System.Windows.Forms.Label();
			this.txtQuery = new System.Windows.Forms.TextBox();
			this.btnEjecutar = new System.Windows.Forms.Button();
			this.txtConectionString = new System.Windows.Forms.TextBox();
			this.lblConectionString = new System.Windows.Forms.Label();
			this.lblResultado = new System.Windows.Forms.Label();
			this.txtResult = new System.Windows.Forms.TextBox();
			this.chkCalculaNombreTabla = new System.Windows.Forms.CheckBox();
			this.txtNombreTabla = new System.Windows.Forms.TextBox();
			this.lblNombreTabla = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// lblQuery
			// 
			this.lblQuery.Location = new System.Drawing.Point(16, 83);
			this.lblQuery.Name = "lblQuery";
			this.lblQuery.TabIndex = 0;
			this.lblQuery.Text = "Query Select";
			// 
			// txtQuery
			// 
			this.txtQuery.Location = new System.Drawing.Point(16, 118);
			this.txtQuery.Multiline = true;
			this.txtQuery.Name = "txtQuery";
			this.txtQuery.Size = new System.Drawing.Size(768, 58);
			this.txtQuery.TabIndex = 2;
			this.txtQuery.Text = "";
			// 
			// btnEjecutar
			// 
			this.btnEjecutar.Location = new System.Drawing.Point(88, 192);
			this.btnEjecutar.Name = "btnEjecutar";
			this.btnEjecutar.Size = new System.Drawing.Size(320, 23);
			this.btnEjecutar.TabIndex = 3;
			this.btnEjecutar.Text = "&Ejecutar";
			this.btnEjecutar.Click += new System.EventHandler(this.btnEjecutar_Click);
			// 
			// txtConectionString
			// 
			this.txtConectionString.Location = new System.Drawing.Point(16, 51);
			this.txtConectionString.Name = "txtConectionString";
			this.txtConectionString.Size = new System.Drawing.Size(768, 20);
			this.txtConectionString.TabIndex = 1;
			this.txtConectionString.Text = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=IndicadoresBolsaDB;Data Source=(localdb)\\MSSQLLocalDB";
			// 
			// lblConectionString
			// 
			this.lblConectionString.Location = new System.Drawing.Point(16, 16);
			this.lblConectionString.Name = "lblConectionString";
			this.lblConectionString.TabIndex = 0;
			this.lblConectionString.Text = "Conection String";
			// 
			// lblResultado
			// 
			this.lblResultado.Location = new System.Drawing.Point(16, 240);
			this.lblResultado.Name = "lblResultado";
			this.lblResultado.TabIndex = 6;
			this.lblResultado.Text = "Resultado";
			// 
			// txtResult
			// 
			this.txtResult.Location = new System.Drawing.Point(16, 272);
			this.txtResult.Multiline = true;
			this.txtResult.Name = "txtResult";
			this.txtResult.ScrollBars = System.Windows.Forms.ScrollBars.Both;
			this.txtResult.Size = new System.Drawing.Size(880, 368);
			this.txtResult.TabIndex = 5;
			this.txtResult.Text = "";
			// 
			// chkCalculaNombreTabla
			// 
			this.chkCalculaNombreTabla.Location = new System.Drawing.Point(432, 192);
			this.chkCalculaNombreTabla.Name = "chkCalculaNombreTabla";
			this.chkCalculaNombreTabla.Size = new System.Drawing.Size(136, 24);
			this.chkCalculaNombreTabla.TabIndex = 4;
			this.chkCalculaNombreTabla.Text = "Calcula nombre tabla";
			// 
			// txtNombreTabla
			// 
			this.txtNombreTabla.Location = new System.Drawing.Point(656, 192);
			this.txtNombreTabla.Name = "txtNombreTabla";
			this.txtNombreTabla.TabIndex = 7;
			this.txtNombreTabla.Text = "";
			// 
			// lblNombreTabla
			// 
			this.lblNombreTabla.Location = new System.Drawing.Point(584, 200);
			this.lblNombreTabla.Name = "lblNombreTabla";
			this.lblNombreTabla.Size = new System.Drawing.Size(72, 23);
			this.lblNombreTabla.TabIndex = 8;
			this.lblNombreTabla.Text = "Nombre tabla";
			// 
			// frmMain
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(920, 662);
			this.Controls.Add(this.lblNombreTabla);
			this.Controls.Add(this.txtNombreTabla);
			this.Controls.Add(this.chkCalculaNombreTabla);
			this.Controls.Add(this.txtResult);
			this.Controls.Add(this.txtConectionString);
			this.Controls.Add(this.txtQuery);
			this.Controls.Add(this.lblResultado);
			this.Controls.Add(this.lblConectionString);
			this.Controls.Add(this.btnEjecutar);
			this.Controls.Add(this.lblQuery);
			this.Name = "frmMain";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Generación de Querys para Inserción";
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// Punto de entrada principal de la aplicación.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new frmMain());
		}

		private void btnEjecutar_Click(object sender, System.EventArgs e)		
		{
			if ((txtConectionString.Text != "")&& (txtQuery.Text != ""))
			{
				txtResult.Text = "";

				try
				{
					txtResult.Text = MontaQuerys(txtConectionString.Text,txtQuery.Text);
				}
				catch(System.Exception ex)
				{
					txtResult.Text = ex.Message.ToString();
				}
			}
			else
				txtResult.Text = "Por favor, compruebe si la cadena de conexión y la query tienen valores establecidos";
		}

		private string MontaQuerys(string strCadenaConexíon, string strQuerySelect)
		{
			const string COMA = ", ";
			string strInsert;
			string strValues;
			string strNombreTabla = nombreTabla(strQuerySelect.ToUpper());
			string strResult = "";
			SqlConnection conn = new SqlConnection(strCadenaConexíon);
			SqlCommand com = new SqlCommand(strQuerySelect,conn);

			conn.Open();
			SqlDataReader dr = com.ExecuteReader();
			
			while (dr.Read())
			{
				int numColumnas = dr.FieldCount;
				strInsert = "INSERT INTO " + strNombreTabla + "( ";
				strValues = "VALUES (";

				for (int i = 0; i < numColumnas; i++)
				{
					strInsert += dr.GetName(i) + COMA;					

					//nulos
					if (dr.GetValue(i) == null) 
						strValues += "null" + COMA;
					else
					{						
						//comilla para cadenas
						if ((dr.GetValue(i).GetType().ToString() == "System.String") 
							|| (dr.GetValue(i).GetType().ToString() == "System.DateTime")
							|| (dr.GetValue(i).GetType().ToString() == "System.Decimal"))
							strValues += "'" +  NoNulo(dr.GetValue(i)) + "'" + COMA;
						else
							strValues += NoNulo(dr.GetValue(i)) + COMA;						
					}
				}

				strInsert = strInsert.Substring(0,strInsert.Length-2);	//quita coma final
				strInsert += ") ";

				strValues = strValues.Substring(0,strValues.Length-2);	//quita coma final
				strValues += ") ";
				
				//Montaje de querys				
				strResult += strInsert + strValues + Convert.ToChar(13) + Convert.ToChar(10);

			}
			dr.Close();
			conn.Close();

			return strResult;
		}

		private string NoNulo(object valor)
		{
            string result = null;

			if (valor.GetType().ToString() == "System.DBNull")
                result = "null";
			else
				if (valor.GetType().ToString() == "System.Decimal")
                    result = valor.ToString().Replace(",",".");
				else
					result = valor.ToString();

            return result.Trim();
		}

		private string nombreTabla(string strQuerySelect)
		{
			if (chkCalculaNombreTabla.Checked)
			{
				int posIni = strQuerySelect.IndexOf("FROM ") +5;
				return strQuerySelect.Substring(posIni,8);  //el nombre de la tabla SIEMPRE ocupa 8 caracteres
			}
			else
				return txtNombreTabla.Text;
		}

	}
}
