using Microsoft.AnalysisServices.Tabular;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp3
{
    class Program
    {
        public static void AddNewRelationShip()
        {
            string serverName = ConfigurationManager.AppSettings["serverName"];
            string cubeKey = ConfigurationManager.AppSettings["cubeKeyName"];

            string serverConnectionString = string.Format("Provider=MSOLAP;Data Source={0}", serverName);

            Server server = new Server();
            server.Connect(serverConnectionString);

            foreach (Database dataBase in server.Databases)
            {
                if (dataBase.ID.Contains(cubeKey))
                {
                    try
                    {
                        var relation = dataBase.Model.Relationships.Where(r => r.Name == "9a3a79d5-555f-457b-9547-e05f2b95f333").FirstOrDefault();
                        if (relation == null)
                        {
                            dataBase.Model.Relationships.Add(new SingleColumnRelationship
                            {
                                Name = "9a3a79d5-555f-457b-9547-e05f2b95f333",
                                FromColumn = dataBase.Model.Tables["EventIndicators"].Columns["OrganizerID"],
                                FromCardinality = RelationshipEndCardinality.Many,
                                ToColumn = dataBase.Model.Tables["DimOrganizer"].Columns["EmployeeId"],
                                ToCardinality = RelationshipEndCardinality.One,
                                CrossFilteringBehavior = CrossFilteringBehavior.OneDirection,
                                IsActive = true
                            });
                            dataBase.Model.SaveChanges();
                            Console.WriteLine(String.Format("{0}-{1}", dataBase.ID,"Update Ok"));
                        }

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(String.Format("{0}-{1}", dataBase.ID, "Update Ko"));
                        Console.WriteLine(ex.Message);
                    }
                }
            }

        }
        public static void AddOrChangeMeasure()
        {
            string serverName = @" localhost\tabular";
            string databaseName = "Contoso";
            string tableName = "Sales";
            string measureName = "Total Sales";
            string measureExpression = "SUMX ( Sales, Sales[Quantity] * Sales[Net Price] )";

            string serverConnectionString = string.Format("Provider=MSOLAP;Data Source={0}", serverName);

            Server server = new Server();
            server.Connect(serverConnectionString);

            Database db = server.Databases[databaseName];
            Model model = db.Model;
            Table table = model.Tables[tableName];
            Measure measure = table.Measures.Find(measureName);
            if (measure == null)
            {
                measure = new Measure() { Name = measureName };
                table.Measures.Add(measure);
            }
            measure.Expression = measureExpression;
            model.SaveChanges();
        }
        static void Main(string[] args)
        {
            AddNewRelationShip();
        }
    }
}
