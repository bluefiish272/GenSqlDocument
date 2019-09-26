using Dapper;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace SqlDocument.Models
{
    public class DbConnection
    {
        private string connectionString = @"Persist Security Info=False;Integrated Security=SSPI;  
    database=SqlDocument;server=MSI";


        private string getSpDescription = @"Select 
                                            [SpName] = obj.name,
                                            [CreateDate] = obj.create_date,
                                            [ModifyDate] = obj.modify_date,
                                            [Description] = mods.definition
                                            From sys.objects AS obj
                                            left join sys.sql_modules AS mods
                                            on obj.object_id = mods.object_id
                                            where obj.type = 'p'";

        public List<SpInfo> GetSpDescription()
        {
            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                List<SpInfo> spInfos = connection.Query<SpInfo>(getSpDescription).ToList();
                return TransferSpInfos(spInfos);
            }
        }
        private List<SpInfo> TransferSpInfos(List<SpInfo> spInfos)
        {
            foreach (var info in spInfos)
            {
                List<string> temp = info.Description.Split("CREATE PROCEDURE").ToList();
                info.Description = temp.FirstOrDefault();
                info.Script = temp.LastOrDefault() == null ? null : $"CREATE PROCEDURE {temp.LastOrDefault()}";
            }
            return spInfos;
        }
    }
}
