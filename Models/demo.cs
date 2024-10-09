using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.ComponentModel.DataAnnotations;


namespace WebApplication1.Models
{
public class demo
    {

        [Required(ErrorMessage = "Division is required.")]
        public string Division { get; set; }

        [Required(ErrorMessage = "Department is required.")]
        public string Department { get; set; }

        [Required(ErrorMessage = "Location is required.")]
        public string Location { get; set; }

        [Required(ErrorMessage = "Document No. is required.")]
        public string DocumentNo { get; set; }

        [Required(ErrorMessage = "Document Name is required.")]
        public string DocumentName { get; set; }

        [Required(ErrorMessage = "Standards are required.")]
        public string Standards { get; set; }

        public string Applicableclauses { get;  set; }
        [Required(ErrorMessage = "Issue No. is required.")]

        public string Issueno { get;  set; }
        [Required(ErrorMessage = "Revision No. is required.")]

        public string Revno { get;  set; }
        [Required(ErrorMessage = "Preparer is required.")]
        public string Preparer { get; set; }

        [Required(ErrorMessage = "Reviewer is required.")]
        public string Reviewer { get; set; }

        [Required(ErrorMessage = "Approver is required.")]
        public string Approver { get; set; }

        [Required(ErrorMessage = "Date is required.")]
        [DataType(DataType.Date)]
        public DateTime SelectedDate { get; set; }

        [Required(ErrorMessage = "Date is required.")]
        [DataType(DataType.Date)]
        public DateTime Revdate { get; set; } 

        [Required(ErrorMessage = "Date is required.")]
        [DataType(DataType.Date)]
        public DateTime Docdate { get; set; }

        [Required(ErrorMessage = "DocumentLevel is required.")]
        public string DocumentLevel { get; set; }

        public string FilePath { get; set; }
        [Required(ErrorMessage = "ID is required.")]
        public int ID { get; set; }


    }

}