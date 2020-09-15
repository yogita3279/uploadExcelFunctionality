
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;
using WebApplication4.Models;

namespace WebApplication4.Models
{
    public class DemoContext:DbContext
    {
        public DemoContext() : base("ConString")
        {


        }
        public DbSet<STARS_SubmittedRouteData> STARS_SubmittedRouteData { get; set; }
    }
}