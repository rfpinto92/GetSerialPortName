using System;
using System.IO;
using System.Linq;
using System.Management;
using System.Threading;
using System.Xml.Linq;

namespace GetSerialPortName
{
    public class GetSerialPortName
    {
        static void Main(string[] args)
        {
            try
            {
                string strExeFilePath = System.Reflection.Assembly.GetExecutingAssembly().Location;// get location of EXE file
                string strWorkPath = System.IO.Path.GetDirectoryName(strExeFilePath); // get location of EXE file without exe file name on the string
                string strSettingsXmlFilePath = System.IO.Path.Combine(strWorkPath, "FTDIs.xml"); // insert FTDIs.xml on the string

                var xml = XDocument.Load(strSettingsXmlFilePath);
                // Query the data and write out a subset of contacts
                var query = from c in xml.Root.Descendants("FTDI")
                            select c.Element("Name").Value;

                string COMPortResult = "-1";

                foreach (string name in query)
                {
                    string _name = name.Replace("\\", "\\\\");

                    COMPortResult = Get_FTDIBUSBSerialPort_Name(_name);

                    if (COMPortResult != "-1")
                    {
                        Console.WriteLine(COMPortResult);
                        return;
                    }

                }
                Console.WriteLine(COMPortResult);
                return;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }
        }

        static string Get_FTDIBUSBSerialPort_Name(string FTDIName)
        {
            string COMPortName = "-1";

            try
            {
                ManagementObjectSearcher searcher =
                   new ManagementObjectSearcher("root\\CIMV2",
                   "SELECT * FROM Win32_PnPEntity WHERE DeviceID = '" + FTDIName + "'"); // Query to find id of FTDI "device instance path"

                //fet all FTDIs and get the com port name
                foreach (ManagementObject queryObj in searcher.Get())
                {
                    COMPortName = queryObj["Name"].ToString();
                    COMPortName = COMPortName.Substring(COMPortName.LastIndexOf('C'), (COMPortName.Length - COMPortName.LastIndexOf('C') - 1));
                }

                return COMPortName;

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return e.Message;
            }
        }





    }
}

