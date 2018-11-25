using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
//using System.Web.UI;
//using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;
//using System.Web.Services;


namespace ConsoleApp3
{
    class Program
    {
   
    static void Main(string[] args)
    {

        DataTable dTable = new DataTable();
        dTable.Columns.Add("ID_QUESTIONARIO", typeof(int));
        dTable.Columns.Add("DS_QUESTIONARIO", typeof(string));
        dTable.Columns.Add("ID_QUESTIONARIO_TIPO", typeof(int));
        dTable.Columns.Add("DT_INICIO", typeof(DateTime));
        dTable.Columns.Add("ID_CLIENTE", typeof(int));
        dTable.Columns.Add("DT_UPDATE", typeof(DateTime));
        dTable.Columns.Add("FL_ATIVO", typeof(char));



        dTable.Rows.Add("1", "Pesquisa Satisfação", "2", "2014-04-14 00:00:00.000", "23", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("2", "Treinamento", "1", "2014-04-14 00:00:00.000", "23", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("3", "Treinamento", "2", "2014-04-14 00:00:00.000", "23", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("4", "Pesquisa Satisfação TORRENT", "2", "2014-04-23 00:00:00.000", "23", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("5", "Treinamento TORRENT", "1", "2014-04-23 00:00:00.000", "23", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("7", "AVALIAÇÃO RESTIVA MARÇO", "1", "2014-06-06 00:00:00.000", "45", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("9", "Avaliação – Linha Ômega", "1", "2014-08-07 00:00:00.000", "40", "2017-04-10 16:23:55.357", "0");
        dTable.Rows.Add("10", "Homolog - Avaliação Omega Migrainesx", "1", "2014-07-31 00:00:00.000", "40", "2017-04-10 16:23:55.367", "0");
        dTable.Rows.Add("11", "Pesquisa de Satisfação - Reunião Nacional de Vendas - Mavsa", "2", "2014-08-01 00:00:00.000", "45", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("12", "Restiva - Revisão Módulo 1", "1", "2014-08-18 00:00:00.000", "45", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("13", "Restiva - Revisão Módulo 2", "1", "2014-08-25 00:00:00.000", "45", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("14", "Restiva - Avaliação Módulo 3", "1", "2014-09-08 00:00:00.000", "45", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("15", "Treinamento Demonstração", "1", "2014-08-25 00:00:00.000", "23", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("16", "Pesquisa Demonstração", "2", "2014-08-25 00:00:00.000", "23", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("17", "Demo - Treinamento", "1", "2014-09-01 00:00:00.000", "23", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("18", "Demo - Pesquisa", "2", "2014-09-01 00:00:00.000", "23", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("19", "Avaliação Teste", "1", "2014-09-03 00:00:00.000", "40", "2017-01-24 10:18:14.747", "1");
        dTable.Rows.Add("20", "Restiva - Reunião de Ciclo -Set/2014", "1", "2014-09-10 00:00:00.000", "45", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("21", "OxyContin - Linha Onco - Reunião de Ciclo", "1", "2014-09-11 00:00:00.000", "45", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("22", "Avaliação Estratégia Ciclo 08 - Gama 1", "1", "2014-09-05 00:00:00.000", "40", "2017-04-10 16:23:55.367", "0");
        dTable.Rows.Add("23", "Avaliação Estratégia Ciclo 08 - Gama 2", "1", "2014-09-05 00:00:00.000", "40", "2017-04-10 16:23:55.370", "0");
        dTable.Rows.Add("26", "Avaliação Estratégia Ciclo 08 - GDs", "1", "2014-09-05 00:00:00.000", "40", "2017-04-10 16:23:55.370", "0");
        dTable.Rows.Add("27", "Teste", "1", "2014-09-05 00:00:00.000", "40", "2017-04-10 16:23:55.370", "0");
        dTable.Rows.Add("28", "Treinamento Citologia/Histologia", "1", "2014-09-05 00:00:00.000", "40", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("30", "OxyContin- Linha Dor e Institucional - Reunião de Ciclo", "1", "2014-09-11 00:00:00.000", "45", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("31", "Treinamento Essencial Celestone", "1", "2014-09-05 00:00:00.000", "40", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("34", "Treinamento Essencial Predsim", "1", "2014-09-05 00:00:00.000", "40", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("35", "Treinamento Essencial Maxsulid", "1", "2014-09-05 00:00:00.000", "40", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("36", "Treinamento Essencial Addera", "1", "2014-09-05 00:00:00.000", "40", "2016-09-19 10:42:43.490", "1");
        dTable.Rows.Add("37", "Treinamento Farmacologia", "1", "2014-09-07 00:00:00.000", "40", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("38", "Treinamento Essencial Dermatologia Básica", "1", "2014-09-08 00:00:00.000", "40", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("39", "Treinamento Essencial Envelhecimento", "1", "2014-09-08 00:00:00.000", "40", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("40", "Treinamento Essencial C-kaderm", "1", "2014-09-08 00:00:00.000", "40", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("41", "Demo - Treinamento", "1", "2014-09-19 00:00:00.000", "23", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("42", "Pesquisa Demonstração", "2", "2014-10-07 00:00:00.000", "23", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("43", "RECICLAGEM GDS MISTOS - 15/10", "1", "2014-10-15 00:00:00.000", "40", "2017-04-10 16:23:55.370", "0");
        dTable.Rows.Add("44", "RECICLAGEM GDS MISTOS - 16/10", "1", "2014-10-16 00:00:00.000", "40", "2017-04-10 16:23:55.370", "0");
        dTable.Rows.Add("46", "RECICLAGEM GDS MISTOS 17/10MANHÃ", "1", "2014-10-17 00:00:00.000", "40", "2017-04-10 16:23:55.373", "0");
        dTable.Rows.Add("47", "RECICLAGEM GDS MISTOS 17/10TARDE", "1", "2014-10-17 00:00:00.000", "40", "2017-04-10 16:23:55.373", "0");
        dTable.Rows.Add("49", "RECICLAGEM GDS MISTOS 17/10MANHÃ", "1", "2014-10-17 00:00:00.000", "40", "2017-04-10 16:23:55.373", "0");
        dTable.Rows.Add("51", "AVALIAÇÃO DE REAÇÃO - RECICLAGEM GDS MISTOS", "2", "2014-10-20 00:00:00.000", "40", "2017-04-10 16:23:55.373", "0");
        dTable.Rows.Add("52", "AVALIAÇÃO DE REAÇÃO - RECICLAGEM GDS MISTOS", "2", "2014-10-20 00:00:00.000", "40", "2017-04-10 16:23:55.377", "0");
        dTable.Rows.Add("53", "Questionário 1 - Prova Não Aleatória", "1", "2014-10-23 00:00:00.000", "53", "2017-09-14 16:31:44.347", "1");
        dTable.Rows.Add("54", "Questionário 2 - Prova Aleatória", "1", "2014-10-23 00:00:00.000", "53", "2017-09-14 16:31:44.347", "1");
        dTable.Rows.Add("55", "Questionário Prova Descritiva", "1", "2014-10-24 00:00:00.000", "53", "2017-09-14 16:31:44.347", "1");
        dTable.Rows.Add("56", "Questionário Pesquisa ", "2", "2014-10-24 00:00:00.000", "53", "2017-09-14 16:31:44.347", "1");
        dTable.Rows.Add("59", "Treinamento Essencial Apraz", "1", "2014-11-10 00:00:00.000", "40", "2017-04-10 16:23:55.377", "0");
        dTable.Rows.Add("60", "Treinamento Essencial Diprospan", "1", "2014-11-10 00:00:00.000", "40", "2017-04-10 16:23:55.380", "0");
        dTable.Rows.Add("61", "Treinamento Essencial Lioram", "1", "2014-11-11 00:00:00.000", "40", "2017-04-10 16:23:55.383", "0");
        dTable.Rows.Add("62", "Treinamento Essencial Lipanon/Cibrato", "1", "2014-11-12 00:00:00.000", "40", "2017-04-10 16:23:55.383", "0");
        dTable.Rows.Add("63", "Treinamento Essencial Lopigrel", "1", "2014-11-12 00:00:00.000", "40", "2017-04-10 16:23:55.383", "0");
        dTable.Rows.Add("64", "Treinamento Essencial Milgamma", "1", "2014-11-12 00:00:00.000", "40", "2017-04-10 16:23:55.383", "0");
        dTable.Rows.Add("65", "Treinamento Essencial Mioflex-A", "1", "2014-11-12 00:00:00.000", "40", "2017-04-10 16:23:55.387", "0");
        dTable.Rows.Add("66", "Treinamento Essencial Scaflam", "1", "2014-11-12 00:00:00.000", "40", "2017-04-10 16:23:55.387", "0");
        dTable.Rows.Add("67", "Questionario de Teste SD", "2", "2014-11-01 00:00:00.000", "53", "2017-09-14 16:31:44.347", "1");
        dTable.Rows.Add("68", "Treinamento Essencial Lopigrel", "1", "2014-11-18 00:00:00.000", "40", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("69", "Treinamento Essencial Milgamma", "1", "2014-11-24 00:00:00.000", "40", "2017-04-10 16:23:55.387", "0");
        dTable.Rows.Add("70", "Teste - Diagnóstico de Conhecimento", "1", "2014-11-26 00:00:00.000", "40", "2017-04-10 16:23:55.387", "0");
        dTable.Rows.Add("71", "Teste 2 - diagnóstico de conhecimento ", "1", "2014-11-26 00:00:00.000", "40", "2017-04-10 16:23:55.390", "0");
        dTable.Rows.Add("72", "Teste 3 - diagnostico de conhecimento", "1", "2014-11-26 00:00:00.000", "40", "2017-04-10 16:23:55.390", "0");
        dTable.Rows.Add("73", "Diagnóstico de Conhecimento - ALFA", "1", "2014-11-28 00:00:00.000", "40", "2017-04-10 16:23:55.390", "0");
        dTable.Rows.Add("74", "Diagnóstico de Conhecimento - BETA", "1", "2014-11-28 00:00:00.000", "40", "2017-04-10 16:23:55.390", "0");
        dTable.Rows.Add("75", "Diagnóstico de Conhecimento - DELTA", "1", "2014-11-28 00:00:00.000", "40", "2017-04-10 16:23:55.390", "0");
        dTable.Rows.Add("76", "Diagnóstico de Conhecimento - ÔMEGA", "1", "2014-11-28 00:00:00.000", "40", "2017-04-10 16:23:55.393", "0");
        dTable.Rows.Add("77", "Diagnóstico de Conhecimento - GAMA 1", "1", "2014-11-28 00:00:00.000", "40", "2017-04-10 16:23:55.393", "0");
        dTable.Rows.Add("78", "Diagnóstico de Conhecimento - GAMA 2", "1", "2014-11-28 00:00:00.000", "40", "2017-04-10 16:23:55.393", "0");
        dTable.Rows.Add("79", "Diagnóstico de Conhecimento - GD SKINCARE", "1", "2014-11-28 00:00:00.000", "40", "2017-04-10 16:23:55.393", "0");
        dTable.Rows.Add("80", "TREINAMENTO EXPRESS - TAMARINE", "1", "2014-12-10 00:00:00.000", "31", "2017-04-10 15:14:48.797", "0");
        dTable.Rows.Add("81", "Teste de Conhecimento: Estomazil Sache, Estomazil Pastilhas, Epocler e Engov", "1", "2015-01-07 00:00:00.000", "31", "2017-04-10 15:14:48.813", "0");
        dTable.Rows.Add("82", "Treinamento Essencial Alivium", "1", "2014-12-17 00:00:00.000", "40", "2017-04-10 16:23:55.393", "0");
        dTable.Rows.Add("83", "Treinamento Essencial Bambair", "1", "2014-12-17 00:00:00.000", "40", "2017-04-10 16:23:55.397", "0");
        dTable.Rows.Add("84", "Treinamento Essencial Digedrat", "1", "2014-12-17 00:00:00.000", "40", "2017-04-10 16:23:55.397", "0");
        dTable.Rows.Add("85", "Treinamento Essencial Fluir", "1", "2014-12-17 00:00:00.000", "40", "2017-04-10 16:23:55.397", "0");
        dTable.Rows.Add("86", "Treinamento Essencial Lisador ", "1", "2014-12-17 00:00:00.000", "40", "2017-04-10 16:23:55.397", "0");
        dTable.Rows.Add("88", "Treinamento Essencial Paxoral", "1", "2014-12-17 00:00:00.000", "40", "2017-04-10 16:23:55.400", "0");
        dTable.Rows.Add("89", "Treinamento Essencial Peridal", "1", "2014-12-17 00:00:00.000", "40", "2017-04-10 16:23:55.400", "0");
        dTable.Rows.Add("90", "Treinamento Essencial Pratium", "1", "2014-12-17 00:00:00.000", "40", "2017-04-10 16:23:55.400", "0");
        dTable.Rows.Add("91", "Treinamento Essencial Rinosoro", "1", "2014-12-17 00:00:00.000", "40", "2017-04-10 16:23:55.400", "0");
        dTable.Rows.Add("92", "Treinamento Essencial Sulbamox BD", "1", "2014-12-17 00:00:00.000", "40", "2017-04-10 16:23:55.400", "0");
        dTable.Rows.Add("93", "Treinamento Essencial Tamarine", "1", "2014-12-17 00:00:00.000", "40", "2017-04-10 16:23:55.413", "0");
        dTable.Rows.Add("94", "Treinamento Essencial Migrainex", "1", "2014-12-17 00:00:00.000", "40", "2017-04-10 16:23:55.413", "0");
        dTable.Rows.Add("95", "Treinamento Essencial Oximax", "1", "2014-12-18 00:00:00.000", "40", "2017-04-10 16:23:55.413", "0");
        dTable.Rows.Add("96", "Treinamento Essencial Pratium", "1", "2014-12-22 00:00:00.000", "40", "2017-04-10 16:23:55.413", "0");
        dTable.Rows.Add("97", "Avaliação Linha Dor - Convenção 2015", "1", "2015-01-20 00:00:00.000", "45", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("98", "Avaliação Linha Onco e Institucional - Convenção 2015", "1", "2015-01-20 00:00:00.000", "45", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("99", "AVALIAÇÃO ITINERES 1", "1", "2015-01-12 00:00:00.000", "40", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("100", "AVALIAÇÃO ITINERES 2", "1", "2015-01-12 00:00:00.000", "40", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("101", "AVALIAÇÃO ITINERES 4", "1", "2015-01-13 00:00:00.000", "40", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("102", "AVALIAÇÃO ITINERES 3", "1", "2015-01-13 00:00:00.000", "40", "2015-11-17 16:22:57.353", "1");
        dTable.Rows.Add("103", "Treinamento Essencial Celestamine", "1", "2015-01-17 00:00:00.000", "40", "2017-04-10 16:23:55.417", "0");
        dTable.Rows.Add("104", "Treinamento Essencial Gingilone", "1", "2015-01-18 00:00:00.000", "40", "2017-04-10 16:23:55.417", "0");
        dTable.Rows.Add("105", "Treinamento Essencial Lucretin", "1", "2015-01-18 00:00:00.000", "40", "2017-04-10 16:23:55.417", "0");
        dTable.Rows.Add("106", "Treinamento Essencial Polaramine", "1", "2015-01-18 00:00:00.000", "40", "2017-04-10 16:23:55.417", "0");
        dTable.Rows.Add("107", "Treinamento Essencial Quadriderm", "1", "2015-01-18 00:00:00.000", "40", "2017-04-10 16:23:55.420", "0");
        dTable.Rows.Add("108", "Treinamento Essencial Blancy", "1", "2015-01-19 00:00:00.000", "40", "2017-04-10 16:23:55.420", "0");
        dTable.Rows.Add("109", "Treinamento Essencial Epidrat Ultra", "1", "2015-01-19 00:00:00.000", "40", "2017-04-10 16:23:55.420", "0");
        dTable.Rows.Add("110", "Treinamento Essencial Hydraporin", "1", "2015-01-19 00:00:00.000", "40", "2017-04-10 16:23:55.420", "0");
        dTable.Rows.Add("111", "Treinamento Essencial Glycare", "1", "2015-01-19 00:00:00.000", "40", "2017-04-10 16:23:55.420", "0");
        dTable.Rows.Add("112", "Treinamento Essencial Ivy C", "1", "2015-01-19 00:00:00.000", "40", "2017-04-10 16:23:55.420", "0");
        dTable.Rows.Add("113", "Treinamento Essencial Perfectha", "1", "2015-01-19 00:00:00.000", "40", "2017-04-10 16:23:55.423", "0");
        dTable.Rows.Add("114", "1. Avaliação das Simuladas - Reunião Regional", "2", "2015-02-09 00:00:00.000", "40", "2017-04-10 16:23:55.423", "0");
        dTable.Rows.Add("115", "2. Avaliação das Simuladas - Reunião Regional", "2", "2015-02-09 00:00:00.000", "40", "2017-04-10 16:23:55.423", "0");
        dTable.Rows.Add("116", "3. Avaliação das Simuladas - Reunião Regional", "2", "2015-02-09 00:00:00.000", "40", "2017-04-10 16:23:55.427", "0");
        dTable.Rows.Add("117", "4. Avaliação das Simuladas - Reunião Regional", "2", "2015-02-09 00:00:00.000", "40", "2017-04-10 16:23:55.427", "0");
        dTable.Rows.Add("118", "5. Avaliação das Simuladas - Reunião Regional", "2", "2015-02-09 00:00:00.000", "40", "2017-04-10 16:23:55.427", "0");
        dTable.Rows.Add("119", "6. Avaliação das Simuladas - Reunião Regional", "2", "2015-02-09 00:00:00.000", "40", "2017-04-10 16:23:55.427", "0");
        dTable.Rows.Add("120", "7. Avaliação das Simuladas - Reunião Regional", "2", "2015-02-09 00:00:00.000", "40", "2017-04-10 16:23:55.430", "0");
        dTable.Rows.Add("121", "8. Avaliação das Simuladas - Reunião Regional", "2", "2015-02-09 00:00:00.000", "40", "2017-04-10 16:23:55.430", "0");
        dTable.Rows.Add("122", "9. Avaliação das Simuladas - Reunião Regional", "2", "2015-02-09 00:00:00.000", "40", "2017-04-10 16:23:55.430", "0");
        dTable.Rows.Add("123", "10. Avaliação das Simuladas - Reunião Regional", "2", "2015-02-09 00:00:00.000", "40", "2017-04-10 16:23:55.430", "0");
        dTable.Rows.Add("124", "11. Avaliação das Simuladas - Reunião Regional", "2", "2015-02-09 00:00:00.000", "40", "2017-04-10 16:23:55.430", "0");
        dTable.Rows.Add("125", "12. Avaliação das Simuladas - Reunião Regional", "2", "2015-02-09 00:00:00.000", "40", "2017-04-10 16:23:55.430", "0");
        dTable.Rows.Add("126", "13. Avaliação das Simuladas - Reunião Regional", "2", "2015-02-09 00:00:00.000", "40", "2017-04-10 16:23:55.433", "0");
        dTable.Rows.Add("127", "14. Avaliação das Simuladas - Reunião Regional", "2", "2015-02-09 00:00:00.000", "40", "2017-04-10 16:23:55.433", "0");
        dTable.Rows.Add("129", "16. Avaliação das Simuladas - Reunião Regional", "2", "2015-02-09 00:00:00.000", "40", "2017-04-10 16:23:55.433", "0");
        dTable.Rows.Add("130", "17. Avaliação das Simuladas - Reunião Regional", "2", "2015-02-09 00:00:00.000", "40", "2017-04-10 16:23:55.437", "0");
        dTable.Rows.Add("131", "18. Avaliação das Simuladas - Reunião Regional", "2", "2015-02-09 00:00:00.000", "40", "2017-04-10 16:23:55.437", "0");
        dTable.Rows.Add("132", "19. Avaliação das Simuladas - Reunião Regional", "2", "2015-02-09 00:00:00.000", "40", "2017-04-10 16:23:55.437", "0");
        dTable.Rows.Add("133", "20. Avaliação das Simuladas - Reunião Regional", "2", "2015-02-09 00:00:00.000", "40", "2017-04-10 16:23:55.437", "0");



        ExcelPackage wkBook = new ExcelPackage();

        //Create the worksheet
        ExcelWorksheet sheet1 = wkBook.Workbook.Worksheets.Add("Teste");



        sheet1.Cells["A1"].LoadFromDataTable(dTable, true);



        //Response.Clear();
        //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        //Response.AddHeader("content-disposition", "attachment;  filename=Tipologia.xlsx");
        //Response.BinaryWrite(wkBook.GetAsByteArray());
        ////Response.AppendCookie(new HttpCookie("DownloadAmostras", "-"));
        //Response.Flush();
        //Response.End();

      while (Console.ReadKey().Key != ConsoleKey.Enter) { }

    }
  }
}

