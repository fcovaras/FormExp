using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace exptblexcel
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Threading.Tasks;

    namespace exptblexcel
    {
        public class IniFile
        {
            public string path;

            [DllImport("kernel32")]
            private static extern long WritePrivateProfileString(string section,
                                                                 string key,
                                                                 string val,
                                                                 string filePath);
            [DllImport("kernel32")]
            private static extern int GetPrivateProfileString(string section,
                                                              string key,
                                                              string def,
                                                              StringBuilder retVal,
                                                              int size,
                                                              string filePath);

            public string DirSearch(string sDir, string pathFile)
            {
                try
                {
                    foreach (string d in Directory.GetDirectories(sDir))
                    {
                        foreach (string f in Directory.GetFiles(d, pathFile))
                        {
                            return f;
                        }
                        DirSearch(d, pathFile);
                    }
                }
                catch (System.Exception excpt)
                {
                    Console.WriteLine(excpt.Message);
                }

                return String.Empty;
            }

            public string VersionSO()
            {
                string sVersion = string.Empty;
                System.OperatingSystem osInfo = System.Environment.OSVersion;

                switch (osInfo.Platform)
                {
                    // La plataforma es Windows 95, Windows 98, 
                    // Windows 98 Segunda edición, o Windows Me.
                    case System.PlatformID.Win32Windows:

                        switch (osInfo.Version.Minor)
                        {
                            case 0:
                                sVersion = "Windows 95";
                                break;
                            case 10:
                                if (osInfo.Version.Revision.ToString() == "2222A")
                                    sVersion = "Windows 98 Segunda edición";
                                else
                                    sVersion = "Windows 98";
                                break;
                            case 90:
                                sVersion = "Windows Me";
                                break;

                            default:
                                sVersion = string.Format("Old Windows Major={0} Minor={1}", osInfo.Version.Major, osInfo.Version.Minor);
                                break;
                        }
                        break;

                    // La plataforma es Windows NT 3.51, Windows NT 4.0, Windows 2000,
                    // o Windows XP.
                    case System.PlatformID.Win32NT:

                        switch (osInfo.Version.Major)
                        {
                            case 3:
                                sVersion = "Windows NT 3.51";
                                break;
                            case 4:
                                sVersion = "Windows NT 4.0";
                                break;
                            case 5:
                                if (osInfo.Version.Minor == 0)
                                    sVersion = "Windows 2000";
                                else
                                    sVersion = "Windows XP";
                                break;
                            case 6:
                                if (osInfo.Version.Minor == 0)
                                    sVersion = "Windows Server";
                                else if (osInfo.Version.Minor == 1)
                                    sVersion = "Windows 7";
                                else if (osInfo.Version.Minor == 2)
                                    sVersion = "Windows 10";
                                else
                                    sVersion = "Windows NT Mayor=6 Minor=" + osInfo.Version.Minor.ToString();
                                break;

                            default:
                                sVersion = string.Format("Windows Major={0} Minor={1}", osInfo.Version.Major, osInfo.Version.Minor);
                                break;
                        }
                        break;

                    default:
                        sVersion = "osInfo.Platform: " + osInfo.Platform.ToString();
                        break;
                }

                return sVersion;
            }

            /// <summary>
            /// INIFile Constructor.
            /// </summary>
            /// <PARAM name="INIPath"></PARAM>
            public IniFile(string INIPath)
            {
                /*
                DirectoryInfo di = new DirectoryInfo(@"C:\windows");

                foreach (var fi in di.GetFiles(INIPath))
                {
                    path = fi.FullName;
                    break;
                }
                */
                string dirPath = Path.GetTempPath();

                string sistemaOperativo = VersionSO();
                if (sistemaOperativo.Equals("Windows NT 3.51") || sistemaOperativo.Equals("Windows NT 4.0") ||
                    sistemaOperativo.Equals("Windows 2000") || sistemaOperativo.Equals("Windows XP"))
                {

                }
                else if (dirPath.IndexOf(@"\AppData\Local\Temp") > 0)
                {
                    string winperIni = dirPath.Replace(@"Temp\", @"VirtualStore\Windows\" + INIPath);
                    if (!File.Exists(winperIni))
                    {
                        int nPos = dirPath.IndexOf(@"\AppData\Local\Temp");
                        dirPath = dirPath.Substring(0, nPos) + @"\Windows";
                    }
                    else
                    {
                        dirPath = dirPath.Replace(@"Temp\", @"VirtualStore\Windows\");
                    }
                }
                else
                {
                    int nPos = dirPath.IndexOf("AppData");
                    dirPath = dirPath.Substring(0, nPos) + "Windows";
                }
                path = dirPath + @"\" + INIPath;
                if (!File.Exists(path))
                {
                    path = INIPath;
                }
            }

            /// <summary>
            /// Write Data to the INI File
            /// </summary>
            /// <PARAM name="Section"></PARAM>
            /// Section name
            /// <PARAM name="Key"></PARAM>
            /// Key Name
            /// <PARAM name="Value"></PARAM>
            /// Value Name
            public void IniWriteValue(string Section, string Key, string Value)
            {
                WritePrivateProfileString(Section, Key, Value, this.path);
            }

            /// <summary>
            /// Read Data Value From the Ini File
            /// </summary>
            /// <PARAM name="Section"></PARAM>
            /// <PARAM name="Key"></PARAM>
            /// <PARAM name="Path"></PARAM>
            /// <returns></returns>
            public string IniReadValue(string Section, string Key)
            {
                StringBuilder temp = new StringBuilder(255);
                int i = GetPrivateProfileString(Section, Key, "", temp,
                                                255, this.path);
                return temp.ToString();

            }

            public List<string> ReadAllLines()
            {
                var lineas = new List<string>();
                string sLine = "";
                StreamReader objReader = new StreamReader(this.path);
                while (sLine != null)
                {
                    sLine = objReader.ReadLine();
                    if (sLine != null)
                        lineas.Add(sLine);
                }
                objReader.Close();

                return lineas;
            }

            public  string f_auth_desencripta(string pass_encryptada)
            {
                pass_encryptada = pass_encryptada.Trim();

                string ls_password_encriptada = "";
                int[] li_set_desplazamiento = { -3, -4, 8, 14, -91, 51, 71, -157, -17, -69, -1, -101, -41, 7, -19, -23, -31, 43, 78, 13, -11, -21, 34, -51, -214, 117 };
                string ls_set_caracteres = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
                string ls_set_numeros = "0123456789";
                string sCaracter = pass_encryptada.Substring(0, 1);
                int li_posicion = -1;
                string ls_set_a_utilizar = string.Empty;

                int i = 0;
                while (i < pass_encryptada.Length)
                {
                    string ls_caracter = pass_encryptada.Substring(i, 1).Trim();
                    sCaracter = ls_caracter;
                    char nCaracter = sCaracter[0];
                    if (nCaracter >= '0' && nCaracter <= '9')
                    {
                        li_posicion = ls_set_numeros.IndexOf(nCaracter);
                        ls_set_a_utilizar = ls_set_numeros;
                    }
                    else
                    {
                        li_posicion = ls_set_caracteres.IndexOf(nCaracter);
                        ls_set_a_utilizar = ls_set_caracteres;
                    }
                    int li_desp = li_set_desplazamiento[i];
                    int j = 0;
                    while (true)
                    {
                        if (j == Math.Abs(li_desp))
                        {
                            break;
                        }
                        if (li_desp > 0)
                        {
                            if (li_posicion == (ls_set_a_utilizar.Length - 1))
                            {
                                li_posicion = 0;
                            }
                            else
                            {
                                li_posicion++;
                            }
                        }
                        else
                        {
                            if (li_posicion == 0)
                            {
                                li_posicion = ls_set_a_utilizar.Length - 1;
                            }
                            else
                            {
                                li_posicion--;
                            }
                        }
                        j++;
                    }
                    ls_password_encriptada = ls_password_encriptada + ls_set_a_utilizar.Substring(li_posicion, 1);

                    i++;
                }

                return ls_password_encriptada;
            }

            public  string f_auth_encripta(string pass_desencrytada)
            {
                string ls_password_encriptada = "";
                int[] li_set_desplazamiento = { -3, -4, 8, 14, -91, 51, 71, -157, -17, -69, -1, -101, -41, 7, -19, -23, -31, 43, 78, 13, -11, -21, 34, -51, -214, 117 };
                string ls_set_caracteres = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
                string ls_set_numeros = "0123456789";
                int li_posicion = -1;
                string ls_set_a_utilizar = string.Empty;

                for (int k = 0; k < li_set_desplazamiento.Length; k++)
                {
                    li_set_desplazamiento[k] = -li_set_desplazamiento[k];
                }

                int i = 0;
                while (i < pass_desencrytada.Length)
                {
                    string ls_caracter = pass_desencrytada.Substring(i, 1).Trim();
                    string sCaracter = ls_caracter;
                    char nCaracter = sCaracter[0];
                    if (nCaracter >= '0' && nCaracter <= '9')
                    {
                        li_posicion = ls_set_numeros.IndexOf(nCaracter);
                        ls_set_a_utilizar = ls_set_numeros;
                    }
                    else
                    {
                        li_posicion = ls_set_caracteres.IndexOf(nCaracter);
                        ls_set_a_utilizar = ls_set_caracteres;
                    }
                    int li_desp = li_set_desplazamiento[i];
                    int j = 0;
                    while (true)
                    {
                        if (j == Math.Abs(li_desp))
                        {
                            break;
                        }
                        if (li_desp > 0)
                        {
                            if (li_posicion == (ls_set_a_utilizar.Length - 1))
                            {
                                li_posicion = 0;
                            }
                            else
                            {
                                li_posicion++;
                            }
                        }
                        else
                        {
                            if (li_posicion == 0)
                            {
                                li_posicion = ls_set_a_utilizar.Length - 1;
                            }
                            else
                            {
                                li_posicion--;
                            }
                        }
                        j++;
                    }
                    ls_password_encriptada = ls_password_encriptada + ls_set_a_utilizar.Substring(li_posicion, 1);

                    i++;
                }

                return ls_password_encriptada;
            }
        }
    }

}
