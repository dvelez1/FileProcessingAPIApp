using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccess.Models
{
    public class EmployeeModel
    {
        [JsonProperty("employee_id", Required = Required.Always)]
        public string EmployeeId { get; set; }
        
        [JsonProperty("full_name", Required = Required.Always)]
        public string FullName { get; set; }
        
        [JsonProperty("job_title", Required = Required.Always)]
        public string JobTitle { get; set; }
        
        [JsonProperty("department", Required = Required.AllowNull)]
        public string Department { get; set; }
        
        [JsonProperty("business_unit", Required = Required.AllowNull)]
        public string BusinessUnit { get; set; }
       
        [JsonProperty("gender", Required = Required.Always)]
        public string Gender { get; set; }

        [JsonProperty("id", Required = Required.Always)]
        public int Id { get; set; }

        [JsonProperty("transaction_date", Required = Required.Always)]
        public DateTime TransactionDate { get; set; }
    }
}
