using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Text;
using System.Configuration;


public class ConexionCall
{

    public static DataTable SqlDTable(string strQuery)//DEVUELVE UNA TABLA A PARTIR DE UNA CONSULTA
    {
        DataTable tabla = new DataTable();
        SqlConnection sConn = ConexionCall.ConexionSql();
           SqlCommand sqlCmd = new SqlCommand(strQuery, sConn);
        SqlDataAdapter sqlAdap = new SqlDataAdapter(sqlCmd);

        sqlCmd.CommandTimeout = 25 * 60;
        tabla.Clear();
        try
        {
            sConn.Open();
            sqlAdap.Fill(tabla);
            sqlAdap.Dispose();
            sqlCmd.Dispose();
        }
        catch (Exception)
        {
            tabla.Clear();
        }
        finally
        {
            sConn.Close();
            sConn.Dispose();
        }
        return tabla;
    }



    public static DataTable SqlDTable2(string strQuery)//DEVUELVE UNA TABLA A PARTIR DE UNA CONSULTA
    {
        DataTable tabla = new DataTable();
        SqlConnection sConn = ConexionCall.ConexionSql();


        SqlCommand sqlCmd = new SqlCommand(strQuery, sConn);
        SqlDataAdapter sqlAdap = new SqlDataAdapter(sqlCmd);
        sqlCmd.CommandTimeout = 10000;

        tabla.Clear();
        try
        {
            sConn.Open();
            sqlAdap.Fill(tabla);
            sqlAdap.Dispose();
            sqlCmd.Dispose();
        }
        catch (Exception)
        {
            tabla.Clear();
        }
        finally
        {
            sConn.Close();
            sConn.Dispose();
        }
        return tabla;
    }



    public static SqlConnection ConexionSql()//ABRE UNA NUEVA CONEXION CON SQLSERVER
    {

      //  string SqlStrConn = ConfigurationManager.AppSettings.Get("ConnString");

        string SqlStrConn = ConfigurationManager.ConnectionStrings["ConnString"].ConnectionString;
        // SqlStrConn = @"data source = DESARROLLONVO\DESARROLLO2005; initial catalog = FDF_TDR; user id = sa; password = marsopa1070";
        SqlConnection ConectarSql = null;
        try
        {
            ConectarSql = new SqlConnection(SqlStrConn);

        }
        catch (Exception ex)
        {
            //  Anthem.Manager.AddScriptForClientSideEval(string.Format("alert('{0}')", ex.Message));
        }
        finally
        {
        }
        return ConectarSql;
    }

    protected void ConectarSQLServer(string ingresar)
    {
        SqlCommand comando = new SqlCommand(ingresar, ConexionCall.ConexionSql());
        try
        {
            comando.Connection.Open();
            comando.ExecuteNonQuery();

        }
        catch (Exception)
        {
            //  Anthem.Manager.AddScriptForClientSideEval(string.Format("alert('{0}')", ex.Message));
        }
        finally
        {
            comando.Connection.Close();
        }
    }

    public static DataTable GeneraTablaValidacion(string login, string password)
    {
        try
        {
            string strQuery = (@"SELECT * FROM tb_acounts WHERE cuenta ='" + login + "' and password = '" + password + "' and estado = 'True'");
            DataTable datos = ConexionCall.SqlDTable(strQuery);
            return datos;
        }
        catch (Exception ex)
        {
            throw ex;

        }
    }

    public static DataTable BuscarAntecedentes(string rut)
    {
        try
        {
            string query = (@"SELECT rut_usuario,nombres,a_paterno,a_materno,direccion,telefono,mail,foto From usuario WHERE rut_usuario='" + rut + "'");

            DataTable datos = ConexionCall.SqlDTable(query);
            return datos;
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }

    public int Autenticar(string login, string password)//AUTENTICACIÓN DE USUARIO 
    {

        int nivelAdm = 0;
        DataTable datos = ConexionCall.GeneraTablaValidacion(login, password);

        if (datos.Rows.Count > 0)
        {
            nivelAdm = Convert.ToInt32(datos.Rows[0]["tipo"]);
            return nivelAdm;
        }
        else
        {
            return 0;
        }

    }

    public bool ejecutorBase(string sqlQry)
    {

        SqlConnection sConn = ConexionCall.ConexionSql();
        SqlCommand sqlCmd = new SqlCommand(sqlQry, sConn);

        try
        {
            sConn.Open();
            sqlCmd.ExecuteNonQuery();
            return true;
        }

        catch (Exception)
        { return false; }

        finally
        {
            sqlCmd.Dispose();
            sConn.Dispose();
            sConn.Close();
        }

    }

    public string ejecutorBaseString(string sqlQry)
    {

        SqlConnection sConn = ConexionCall.ConexionSql();
        SqlCommand sqlCmd = new SqlCommand(sqlQry, sConn);

        try
        {
            sConn.Open();
            sqlCmd.ExecuteNonQuery();
            return "ok";
        }

        catch (Exception ex)
        { return ex.Message; }

        finally
        {
            sqlCmd.Dispose();
            sConn.Dispose();
            sConn.Close();
        }

    }
    public int plagoenf(string Qry) //REVISA SI ES ENFERMADAD O PLAGA 
    {
        int EoP = 0;
        string val = string.Empty;
        DataTable tabla = new DataTable();
        tabla = ConexionCall.SqlDTable(Qry);

        try
        {
            val = tabla.Rows[0][0].ToString();
            return EoP = int.Parse(val);
        }
        catch (Exception)
        {
            return EoP = 0;
        }
    }

    public bool[] cargaCheck(string sqlQry)
    {
        string[] val = new string[8];

        bool[] retVal = new bool[8];
        DataTable tabla = new DataTable();
        tabla = ConexionCall.SqlDTable(sqlQry);
        val[0] = tabla.Rows[0]["yema"].ToString();
        val[1] = tabla.Rows[0]["boton"].ToString();
        val[2] = tabla.Rows[0]["florA"].ToString();
        val[3] = tabla.Rows[0]["cuajado"].ToString();
        val[4] = tabla.Rows[0]["pinton"].ToString();
        val[5] = tabla.Rows[0]["cosecha"].ToString();
        val[6] = tabla.Rows[0]["poscosecha"].ToString();
        val[7] = tabla.Rows[0]["receso"].ToString();
        for (int i = 0; i < 9; i++)
        {
            if (val[i] == "True")
                retVal[i] = true;
            else
                retVal[i] = false;

        }
        return retVal;
    }

    public bool Verificador(string strQuery)//BUSCA SI  DATOS SE ENCUENTRAN EN LA BD
    {
        DataTable tabla = new DataTable();
        try
        {
            tabla = ConexionCall.SqlDTable(strQuery);

            if (tabla.Rows.Count > 0)
            {

                tabla.Dispose();
                return true;
            }
            else
            {
                return false;
            }
        }
        catch (Exception)
        {
            return false;
        }

    }

    public static string devuelveValor(string strQuery)// solo de vuelve un solo valor
    {

        string val = string.Empty;
        DataTable tabla = new DataTable();
        tabla = ConexionCall.SqlDTable(strQuery);

        try
        {
            return val = tabla.Rows[0][0].ToString();

        }
        catch (Exception)
        {
            return val = "-1";
        }
    }


    public static string devuelveValor2(string strQuery)// solo de vuelve un solo valor
    {

        string val = string.Empty;
        DataTable tabla = new DataTable();
        tabla = ConexionCall.SqlDTable(strQuery);

        try
        {
            return val = tabla.Rows[0][0].ToString();

        }
        catch (Exception)
        {
            return  "--";
        }
    }





    public static int devuelveValorINT(string strQuery)// solo de vuelve un solo valor entero
    {

        int val = 0;
        DataTable tabla = new DataTable();
        tabla = ConexionCall.SqlDTable(strQuery);

        try
        {
            return val = Convert.ToInt32(tabla.Rows[0][0]);
        }
        catch (Exception)
        {
            return val = 0;
        }
    }

}
