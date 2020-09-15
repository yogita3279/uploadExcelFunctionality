using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;

namespace WebApplication4.Models
{
	public class STARS_SubmittedRouteData
	{
		[Key] public int SubmittedRouteDataId { get; set; }

		public int ImportHistoryId { get; set; }

		public string CountyDistrictCode { get; set; }
		public string StateBusNumber { get; set; }

		public string DistrictName { get; set; }
		public string RouteTypeCode { get; set; }
		public string DistrictRouteNumber { get; set; }
		public string DistrictBusNumber { get; set; }
		public int StateRouteNumber { get; set; }
		public int StopNumber { get; set; }
		public double StopLatitude { get; set; }
		public double StopLongitude { get; set; }
		public string StopDescription { get; set; }
		public string DestinationName { get; set; }
		public string DestinationIdentifier { get; set; }
		public string DestinationLatitude { get; set; }
		public string DestinationLongitude { get; set; }
		public int AssignedStudents { get; set; }
	}
}