using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace GoogleSheetExcel
{
    public class ProgramSheets
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json


        public static String Run(string credentials, String spreadsheetId, String range)
        {
            string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
            string ApplicationName = "Google Sheets API .NET Quickstart";

            UserCredential credential;

            using (var stream = new FileStream(credentials, FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets,Scopes,"user",CancellationToken.None,new FileDataStore(credPath, true)).Result;
                //Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,ApplicationName = ApplicationName,
            });

            // Define request parameters.
            //String spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
            //String spreadsheetId = "1i1mrDtkYE0Wn03YkTD8hVVd2T0EZdPnDualX4g19xgc";
            //String range = "Class Data!A2:E";
            ValueRange response = null;
             try { 
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
             response = request.Execute();
             }
             catch (Exception ex)
             {
                 Console.WriteLine("error"+ex.Message.ToString());
                 return "error" + ex.Message.ToString();
             }
            // Prints the names and majors of students in a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit

             
              IList<IList<Object>> values = response.Values;

              List<CamposFile> campo = new List<CamposFile>();

               if (values != null && values.Count > 0)
               {
                   //Console.WriteLine("values.Count: " + values.Count.ToString());
                   //Console.WriteLine("Name, Major");
                   foreach (var row in values)
                   {
                       CamposFile file = new CamposFile();
                       int consecutivo = 0;
                       try {
                           file.Fecha_Salida = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Planta_Origen = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Trailer_Unidad = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Placas = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Dock = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Guia_ILS = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Destino_Final = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Proveedor = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Tipo_de_Unidad = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Trip = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Hora_Solicitud_cliente = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Hora_Correo_enviado_cliente = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Hora_Rec_Progr_cliente = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Hora_Colocacion_unidad = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Load_ID_Amado = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Chofer_Mexicano = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Tractor = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Arribo_Planta = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Sellos = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Despacho_Planta = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Caseta_Hillo = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Querobabi_llegada = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Querobabi_entra_Inspeccion = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Querobabi_Tipo_Inspeccion = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Querobabi_Salida = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Querobabi_Sellos = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Caseta_Magdalena = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Llegada_yarda_Nogales_Sonora = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Facturas_Recibidas = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Entries_Listos = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Pedimento_Listos = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Enviado_INBOND_Ticsa = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Enviado_ILS = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Cotejado_ILS = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Hora_Llegada_Custodios = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Inspeccion_Mecanica_entra = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Inspeccion_Mecanica_sale = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Hora_Operador_cruce_Listo = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Inspeccion_K9_entra = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Inspeccion_K9_sale = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Sellos_despues_K9 = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Despacho_yarda_a_fila = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Operador_de_cruce = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Tractor_de_cruce = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.En_Fila_Recinto = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Aduana_Mexicana_Sello = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Aduana_Mexicana_Entro_recinto = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Aduana_Mexicana_Hr_entrada_inspeccion = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Aduana_Mexicana_tipo_inspeccion = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Aduana_Mexicana_Hr_salida_inspeccion = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Aduana_Mexicana_sello2 = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Aduana_Americana_Entro_recinto = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Aduana_Americana_Sello = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Aduana_Americana_Hr_entrada_inspeccion = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Aduana_Americana_tipo_inspeccion = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Aduana_Americana_Hr_salida_inspeccion = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Aduana_Americana_sello2 = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Aduana_Americana_Hora_Arribo_Vidal = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Responsable_Hora_recepcion = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Responsable_Sello_Recibe_Caja = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Responsable_Hora_descarga = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Responsable_POD = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Responsable_Hora_Despacho = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Responsable_Sello = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Despacho_Salida_Nogales = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Despacho_Sello = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Despacho_NombreOperador = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Despacho_Celular = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Despacho_Tractor = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Hora_Programada_Entrega = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Hora_Real_Arribo_Destino = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Hora_Recibido = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.POD = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Actualizaciones_1 = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Actualizaciones_2 = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Actualizaciones_3 = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Actualizaciones_4 = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Actualizaciones_5 = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Actualizaciones_6 = row[consecutivo] != null ? row[consecutivo].ToString() : "";
                           consecutivo += 1;
                           file.Comentarios = row[consecutivo] != null ? row[consecutivo].ToString() : "";

                       // Print columns A and E, which correspond to indices 0 and 4.
                       //Console.WriteLine("{0}, {1}", row[0], row[4]);
                           }
                       catch (Exception ex)
                       {
                           //Console.WriteLine("ex:" + ex.Message.ToString());
                           
                       }

                       campo.Add(file);

                       //Console.WriteLine("campo.Count: "+campo.Count.ToString());
                   }
               }
               else
               {
                   Console.WriteLine("No data found.");
               }            

              Google.Apis.Json.NewtonsoftJsonSerializer json = new Google.Apis.Json.NewtonsoftJsonSerializer();
              //Google.Apis.Json.NewtonsoftJsonSerializer json = new Google.Apis.Json.NewtonsoftJsonSerializer();


              string result = json.Serialize(campo);//"";

              //Console.WriteLine(result.ToString());
              return result.ToString(); 
        }

    }

    public class CamposFile
    {
        public string Fecha_Salida { get; set; }
        public string Planta_Origen { get; set; }
        public string Trailer_Unidad { get; set; }
        public string Placas { get; set; }
        public string Dock { get; set; }
        public string Guia_ILS { get; set; }
        public string Destino_Final { get; set; }
        public string Proveedor { get; set; }
        public string Tipo_de_Unidad { get; set; }
        public string Trip { get; set; }
        public string Hora_Solicitud_cliente { get; set; }
        public string Hora_Correo_enviado_cliente { get; set; }
        public string Hora_Rec_Progr_cliente { get; set; }
        public string Hora_Colocacion_unidad { get; set; }
        public string Load_ID_Amado { get; set; }
        public string Chofer_Mexicano { get; set; }
        public string Tractor { get; set; }
        public string Arribo_Planta { get; set; }
        public string Sellos { get; set; }
        public string Despacho_Planta { get; set; }
        public string Caseta_Hillo { get; set; }
        public string Querobabi_llegada { get; set; }
        public string Querobabi_entra_Inspeccion { get; set; }
        public string Querobabi_Tipo_Inspeccion { get; set; }
        public string Querobabi_Salida { get; set; }
        public string Querobabi_Sellos { get; set; }
        public string Caseta_Magdalena { get; set; }
        public string Llegada_yarda_Nogales_Sonora { get; set; }
        public string Facturas_Recibidas { get; set; }
        public string Entries_Listos { get; set; }
        public string Pedimento_Listos { get; set; }
        public string Enviado_INBOND_Ticsa { get; set; }
        public string Enviado_ILS { get; set; }
        public string Cotejado_ILS { get; set; }
        public string Hora_Llegada_Custodios { get; set; }
        public string Inspeccion_Mecanica_entra { get; set; }
        public string Inspeccion_Mecanica_sale { get; set; }
        public string Hora_Operador_cruce_Listo { get; set; }
        public string Inspeccion_K9_entra { get; set; }
        public string Inspeccion_K9_sale { get; set; }
        public string Sellos_despues_K9 { get; set; }
        public string Despacho_yarda_a_fila { get; set; }
        public string Operador_de_cruce { get; set; }
        public string Tractor_de_cruce { get; set; }
        public string En_Fila_Recinto { get; set; }
        public string Aduana_Mexicana_Sello { get; set; }
        public string Aduana_Mexicana_Entro_recinto { get; set; }
        public string Aduana_Mexicana_Hr_entrada_inspeccion { get; set; }
        public string Aduana_Mexicana_tipo_inspeccion { get; set; }
        public string Aduana_Mexicana_Hr_salida_inspeccion { get; set; }
        public string Aduana_Mexicana_sello2 { get; set; }
        public string Aduana_Americana_Entro_recinto { get; set; }
        public string Aduana_Americana_Sello { get; set; }
        public string Aduana_Americana_Hr_entrada_inspeccion { get; set; }
        public string Aduana_Americana_tipo_inspeccion { get; set; }
        public string Aduana_Americana_Hr_salida_inspeccion { get; set; }
        public string Aduana_Americana_sello2 { get; set; }
        public string Aduana_Americana_Hora_Arribo_Vidal  { get; set; }
        public string Responsable_Hora_recepcion { get; set; }
        public string Responsable_Sello_Recibe_Caja { get; set; }
        public string Responsable_Hora_descarga { get; set; }
        public string Responsable_POD { get; set; }
        public string Responsable_Hora_Despacho  { get; set; }
        public string Responsable_Sello { get; set; }
        public string Despacho_Salida_Nogales { get; set; }
        public string Despacho_Sello { get; set; }
        public string Despacho_NombreOperador { get; set; }
        public string Despacho_Celular { get; set; }
        public string Despacho_Tractor { get; set; }
        public string Hora_Programada_Entrega { get; set; }
        public string Hora_Real_Arribo_Destino { get; set; }
        public string Hora_Recibido { get; set; }
        public string POD { get; set; }
        public string Actualizaciones_1 { get; set; }
        public string Actualizaciones_2 { get; set; }
        public string Actualizaciones_3 { get; set; }
        public string Actualizaciones_4 { get; set; }
        public string Actualizaciones_5 { get; set; }
        public string Actualizaciones_6 { get; set; }
        public string Comentarios { get; set; }

    }
}
