using OfficeOpenXml;
using System.Data.Common;
using System.Diagnostics;
using System.Drawing;
using System.Reflection.Metadata.Ecma335;
using System.Runtime.ConstrainedExecution;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

//Setter and Getters
public class informacion
{
    public string Mes { get; set; }
    public string ReceptorNacionalidad { get; set; }
    public string NumRegIdTrib { get; set; }
    public string NomDenRazSocR { get; set; }
    public string MesIni { get; set; }
    public string MesFin { get; set; }
    public string Ejerc { get; set; }
    public string montoTotOperacion { get; set; }
    public string montoTotGrav { get; set; }
    public string MmontoTotExentes { get; set; }
    public string montoTotRet { get; set; }
    public string montoRet { get; set; }
    public string TipoPagoRet { get; set; }
    public string EsBenefEfectDelCobro { get; set; }
    public string NoBeneficiario { get; set; }
    public string NoBeneficiariods { get; set; }
    public string Beneficiario { get; set; }

}

class Program
{
    static void Main(string[] args)
    {

        bool exit = false;

        //Primer menu de opciones
        while (!exit)
        {
            Console.Clear();
            Console.WriteLine("=== Selecciona una opcion ===");
            Console.WriteLine("1. Asignar ruta");
            Console.WriteLine("2. Salir");
            Console.WriteLine("=======================");
            Console.Write("Elige una opción: ");

            string input = Console.ReadLine();

            //Acciones de seleccion 
            switch (input)
            {
                case "1":
                    ruta();
                    break;
                case "2":
                    Console.WriteLine("Saliendo del programa...");
                    exit = true;
                    break;
                default:
                    Console.WriteLine("Opción no válida. Por favor, elige una opción válida.");
                    break;
            }

            Console.WriteLine("\nPara intentar nuevamente presione cualquier tecla");
            Console.ReadKey();
        }
    }
    private static void ruta()
    {
        Console.Write("\nIngresa la ruta del archivo: ");
        string filePath = Console.ReadLine();

        // Comprobar archivo es accesible
        if (!File.Exists(filePath))
        {
            Console.WriteLine("Error al leer el archivo.");
            return;
        }

        // Leera el excel
        FileInfo fileInfo = new FileInfo(filePath);

        List<informacion> informaciones = new List<informacion>();

        using (ExcelPackage package = new ExcelPackage(fileInfo))
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Se asigna la hoja con la que se trabajara

            int rowCount = worksheet.Dimension.Rows;
            int columnCount = worksheet.Dimension.Columns;
            int totals = rowCount - 1;

            // Recorremos las filas y extraemos los datos
            for (int row = 2; row <= rowCount; row++) // Comenzamos en la fila 2 para omitir el encabezado
            {
                informacion informacion = new()
                {

                    Mes = worksheet.Cells[row, 1].Value?.ToString(),
                    ReceptorNacionalidad = worksheet.Cells[row, 2].Value?.ToString(),
                    NumRegIdTrib = worksheet.Cells[row, 3].Value?.ToString(),
                    NomDenRazSocR = worksheet.Cells[row, 4].Value?.ToString(),
                    MesIni = worksheet.Cells[row, 5].Value?.ToString(),
                    MesFin = worksheet.Cells[row, 6].Value?.ToString(),
                    Ejerc = worksheet.Cells[row, 7].Value?.ToString(),
                    montoTotOperacion = worksheet.Cells[row, 8].Value?.ToString(),
                    montoTotGrav = worksheet.Cells[row, 9].Value?.ToString(),
                    MmontoTotExentes = worksheet.Cells[row, 10].Value?.ToString(),
                    montoTotRet = worksheet.Cells[row, 11].Value?.ToString(),
                    montoRet = worksheet.Cells[row, 12].Value?.ToString(),
                    TipoPagoRet = worksheet.Cells[row, 13].Value?.ToString(),
                    EsBenefEfectDelCobro = worksheet.Cells[row, 14].Value?.ToString(),
                    NoBeneficiario = worksheet.Cells[row, 15].Value?.ToString(),
                    NoBeneficiariods = worksheet.Cells[row, 17].Value?.ToString(),
                    Beneficiario = worksheet.Cells[row, 19].Value?.ToString()
                };
                informaciones.Add(informacion);
            }

            Console.WriteLine("\n\nSe generaran: " + $"{totals}" + " Archivos xml");
            Console.WriteLine("=== Para continuar selecciona una opcion ===");
            Console.WriteLine("1. Realizar proceso");
            Console.WriteLine("2. Salir");
            Console.WriteLine("\nEscribe tu eleccion: ");

            string input2 = Console.ReadLine();

            switch (input2)
            {
                case "1":
                    Console.Write("Ingresa la ruta de guardado del archivo: ");
                    string ruta = Console.ReadLine();
                    generar(ruta, informaciones);
                    break;
                case "2":
                    Console.WriteLine("Saliendo del programa...");
                    /*exit = true;*/
                    break;
                default:
                    Console.WriteLine("Opción no válida. Por favor, elige una opción válida.");
                    break;
            }

        }

    }

    private static void generar(string? ruta, List<informacion> informaciones)
    {
        foreach (var informacion in informaciones)
        {

           /* var mes = informacion.Mes;
            Console.WriteLine("Mes: " + informacion.Mes);
            Console.WriteLine("Receptor Nacionalidad: " + informacion.ReceptorNacionalidad);
            Console.WriteLine("NumRegIdTrib: " + informacion.NumRegIdTrib);
            Console.WriteLine("NomDenRazSocR: " + informacion.NomDenRazSocR);
            Console.WriteLine("MesIni: " + informacion.MesIni);
            Console.WriteLine("MesFin: " + informacion.MesFin);
            Console.WriteLine("Ejerc: " + informacion.Ejerc);
            Console.WriteLine("montoTotOperacion: " + informacion.montoTotOperacion);
            Console.WriteLine("montoTotGrav: " + informacion.montoTotGrav);
            Console.WriteLine("MmontoTotExentes: " + informacion.MmontoTotExentes);
            Console.WriteLine("montoTotRet: " + informacion.montoTotRet);
            Console.WriteLine("montoRet: " + informacion.montoRet);
            Console.WriteLine("TipoPagoRet: " + informacion.TipoPagoRet);
            Console.WriteLine("EsBenefEfectDelCobro: " + informacion.EsBenefEfectDelCobro);
            Console.WriteLine("NoBeneficiario: " + informacion.NoBeneficiario);
            Console.WriteLine("NoBeneficiariods: " + informacion.NoBeneficiariods);
            Console.WriteLine("Beneficiario: " + informacion.Beneficiario);*/

            //estructura donde ira todo el contenido que se generara
            string ident = "<?xml version=\"1.0\" encoding=\"utf-8\" ?>" +
                "<retenciones:Retenciones xmlns:retenciones=\"http://www.sat.gob.mx/esquemas/retencionpago/1\" " +
                "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:pagosaextranjeros=\"http://www.sat.gob.mx/esquemas/retencionpago/1/pagosaextranjeros\" xsi:schemaLocation=\"http://www.sat.gob.mx/retenciones http://www.sat.gob.mx/esquemas/retencionpago/1/catalogos/retencionpagov1.xsd\" " +
                "Version=\"1.0\" FolioInt=\"EXT29\" Sello=\"\" NumCert=\"00001000000505788565\" Cert=\"\" >" +
            "</retenciones:Retenciones>";

            //colocar el elemento no soportado " : "
            XDocument doc = XDocument.Parse(ident);
            XElement partyTaxScheme = doc.Root;
            XNamespace retenciones = partyTaxScheme.GetNamespaceOfPrefix("retenciones");
            XNamespace pagosaextranjeros = partyTaxScheme.GetNamespaceOfPrefix("pagosaextranjeros");

            //Creara un nuevo nodo con sus respectivos atributos y valores
            XElement Emisor = new XElement(retenciones + "Emisor", new object[] {
                new XAttribute("RFCEmisor", "FSG7712283P5"),
                new XAttribute("NomDenRazSocE", "FUNDACION SANTOS Y DE LA GARZA EVIA IBP"),
                });
            partyTaxScheme.Add(Emisor);

            //Creara un nuevo nodo con sus respectivos atributos y valores
            XElement Receptor = new XElement(retenciones + "Receptor", new object[] {
                new XAttribute("Nacionalidad", informacion.ReceptorNacionalidad),
                });
            partyTaxScheme.Add(Receptor);
            //agregara datos dentro del nodo
            Receptor.Add(new XElement(retenciones + "Extranjero", new object[] {
                new XAttribute("NumRegIdTrib", informacion.NumRegIdTrib),
                new XAttribute("NomDenRazSocR", informacion.NomDenRazSocR)
            }));

            //Creara un nuevo nodo con sus respectivos atributos y valores
            XElement Periodo = new XElement(retenciones + "Periodo", new object[] {
                new XAttribute("MesIni", informacion.MesIni),
                new XAttribute("MesFin", informacion.MesFin),
                new XAttribute("Ejerc", informacion.Ejerc),
                });
            partyTaxScheme.Add(Periodo);

            //Creara un nuevo nodo con sus respectivos atributos y valores
            XElement Totales = new XElement(retenciones + "Totales", new object[] {
                new XAttribute("montoTotOpera"+"cion", informacion.montoTotGrav),
                new XAttribute("montoTotGrav", informacion.montoTotGrav),
                new XAttribute("montoTotExent", informacion.MmontoTotExentes),
                new XAttribute("montoTotRet", informacion.montoTotRet),
                });
            partyTaxScheme.Add(Totales);
            //agregara datos dentro del nodo
            Totales.Add(new XElement(retenciones + "ImpRetenidos", new object[] {
                new XAttribute("montoRet", informacion.montoRet),
                new XAttribute("TipoPagoRet", informacion.TipoPagoRet)
            }));


            //Creara un nuevo nodo con sus respectivos atributos y valores
            XElement Complemento = new XElement(retenciones + "Complemento");
            Complemento.Add(new XElement(pagosaextranjeros + "Pagosextranjeros", new object[] {
                new XAttribute("Version", "1.0"),
                new XAttribute("EsBenefEfectDelCobro", informacion.EsBenefEfectDelCobro)
                }));

            if (informacion.EsBenefEfectDelCobro == "No")
            {
                Complemento.Add(new XElement(pagosaextranjeros + "ImpRetenidos", new object[] {
                        new XAttribute("PaisDeResidParaEfecFisc", informacion.NoBeneficiario),
                        new XAttribute("ConceptoPago", informacion.TipoPagoRet),
                        new XAttribute("DescripcionConcepto", informacion.NoBeneficiariods)
                         }));
            }
           
            partyTaxScheme.Add(Complemento);


            Console.WriteLine("El proceso ha sido ralizado con exito");
            string xmlFilePath = ruta + informacion.montoTotOperacion + ".xml";
            doc.Save(xmlFilePath);

            try
            {
                // Verificamos si la ruta es válida antes de abrir la carpeta.
                if (System.IO.Directory.Exists(ruta))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = ruta,
                        UseShellExecute = true
                    });
                }
                else
                {
                    Console.WriteLine("La carpeta no existe o la ruta no es válida.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al abrir la carpeta: " + ex.Message);
            }

        }
    }
}


