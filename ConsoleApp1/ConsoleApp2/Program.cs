
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;


namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            DataTable dTable = new DataTable();
            dTable.Columns.Add("ID", typeof(int));
            dTable.Columns.Add("NOME", typeof(string));
            dTable.Columns.Add("SOBRENOME", typeof(string));
            
            
            dTable.Rows.Add("1","Rafaela","Lima");
            dTable.Rows.Add("2","Rafaela","Pinheiro");
            dTable.Rows.Add("3","Rafaela","Messina");
            dTable.Rows.Add("4","Rafaela","Laranja");
            dTable.Rows.Add("5","Rafaela","Pera");
            dTable.Rows.Add("6","Rafaela","Pereira");
            dTable.Rows.Add("7","Rafaela","Rodrigues");
            dTable.Rows.Add("8","Rafaela","Fernades");
            
            dTable.Rows.Add("9","Rodrigo","Messina");
            dTable.Rows.Add("10","Rodrigo","Oliveira");
            dTable.Rows.Add("11","Rodrigo","Pinto");
            dTable.Rows.Add("12","Rodrigo","Lima");
            

            //Agrupando
            var query = from row in dTable.AsEnumerable()
                        group row by row.Field<string>("NOME") into grp
                        select new
                        {
                            NOME = grp.Key,
                            item = grp.First()
                        };
            foreach (var grp in query)
            {
                Console.WriteLine(String.Format("NOME-> {0}", grp.NOME));

                foreach (var item in grp.item.ItemArray)
                {
                    Console.WriteLine(String.Format("ITEM -> {0}", item));
                }
            }


            while (Console.ReadKey().Key != ConsoleKey.Enter) { }
        }
    }
}