using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Web;
using System.Web.Mvc;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            List<demo> demoList = new List<demo>();

            string connectionString = ConfigurationManager.ConnectionStrings["MySqlConnection"].ConnectionString;
            string query = "SELECT ID,Division, Department, Location, DocumentNo, DocumentName, Standards, Applicableclauses, Issueno, Revno, Preparer, Reviewer, Approver, SelectedDate, RevDate, DocDate,DocumentLevel FROM documents"; // Ensure columns match the table

            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                using (MySqlCommand cmd = new MySqlCommand(query, conn))
                {
                    conn.Open();
                    using (MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            demo obj = new demo()
                            {
                                ID= (int)reader["ID"],
                                Division = reader["Division"].ToString(),
                                Department = reader["Department"].ToString(),
                                Location = reader["Location"].ToString(),
                                DocumentNo = reader["DocumentNo"].ToString(),
                                DocumentName = reader["DocumentName"].ToString(),
                                Standards = reader["Standards"].ToString(),
                                Applicableclauses = reader["Applicableclauses"].ToString(),
                                Issueno = reader["Issueno"].ToString(),
                                Revno = reader["Revno"].ToString(),
                                Preparer = reader["Preparer"].ToString(),
                                Reviewer = reader["Reviewer"].ToString(),
                                Approver = reader["Approver"].ToString(),
                                SelectedDate = Convert.ToDateTime(reader["SelectedDate"]),
                                Revdate = Convert.ToDateTime(reader["RevDate"]),
                                Docdate = Convert.ToDateTime(reader["DocDate"]),
                                DocumentLevel = reader["DocumentLevel"].ToString(),
                            };
                            demoList.Add(obj);
                        }
                    }
                }
            }

            return View(demoList);
        }



        [HttpGet]
        public ActionResult About()
        {
            ViewBag.Message = "Insert the values";
            return View(new demo());
        }

        [HttpPost]
        public ActionResult About(demo model, HttpPostedFileBase file)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    // Check if a file was uploaded
                    if (file != null && file.ContentLength > 0)
                    {
                        // Generate a unique file name
                        string fileName = Path.GetFileNameWithoutExtension(file.FileName);
                        string extension = Path.GetExtension(file.FileName);
                        string uniqueFileName = fileName + "_" + Guid.NewGuid().ToString() + extension;

                        // Set the path where the file will be saved
                        string path = Path.Combine(Server.MapPath("~/UploadedFiles"), uniqueFileName);

                        // Save the file to the specified path
                        file.SaveAs(path);

                        // Store the file path in the model
                        model.FilePath = "/UploadedFiles/" + uniqueFileName;

                        // Debugging: Check if the file path is set correctly
                        System.Diagnostics.Debug.WriteLine("File Path: " + model.FilePath);
                    }

                    // Get the connection string from Web.config
                    string connectionString = ConfigurationManager.ConnectionStrings["MySqlConnection"].ConnectionString;

                    // Insert query
                    string query = "INSERT INTO documents (Division, Department, Location, DocumentNo, DocumentName, Standards, Applicableclauses, Issueno, Revno, Preparer, Reviewer, Approver, SelectedDate, RevDate, DocDate, DocumentLevel, FilePath) " +
                                   "VALUES (@Division, @Department, @Location, @DocumentNo, @DocumentName, @Standards, @Applicableclauses, @Issueno, @Revno, @Preparer, @Reviewer, @Approver, @SelectedDate, @RevDate, @DocDate, @DocumentLevel, @FilePath)";

                    using (MySqlConnection conn = new MySqlConnection(connectionString))
                    {
                        using (MySqlCommand cmd = new MySqlCommand(query, conn))
                        {
                            // Open the connection
                            conn.Open();

                            // Add parameters to the query
                            cmd.Parameters.AddWithValue("@Division", model.Division);
                            cmd.Parameters.AddWithValue("@Department", model.Department);
                            cmd.Parameters.AddWithValue("@Location", model.Location);
                            cmd.Parameters.AddWithValue("@DocumentNo", model.DocumentNo);
                            cmd.Parameters.AddWithValue("@DocumentName", model.DocumentName);
                            cmd.Parameters.AddWithValue("@Standards", model.Standards);
                            cmd.Parameters.AddWithValue("@Applicableclauses", model.Applicableclauses);
                            cmd.Parameters.AddWithValue("@Issueno", model.Issueno);
                            cmd.Parameters.AddWithValue("@Revno", model.Revno);
                            cmd.Parameters.AddWithValue("@Preparer", model.Preparer);
                            cmd.Parameters.AddWithValue("@Reviewer", model.Reviewer);
                            cmd.Parameters.AddWithValue("@Approver", model.Approver);
                            cmd.Parameters.AddWithValue("@SelectedDate", model.SelectedDate);
                            cmd.Parameters.AddWithValue("@RevDate", model.Revdate);
                            cmd.Parameters.AddWithValue("@DocDate", model.Docdate);
                            cmd.Parameters.AddWithValue("@DocumentLevel", model.DocumentLevel);
                            cmd.Parameters.AddWithValue("@FilePath", model.FilePath);  // Save the file path

                            // Debugging: Check if the query is executing
                            System.Diagnostics.Debug.WriteLine("Executing query...");

                            // Execute the query
                            cmd.ExecuteNonQuery();
                        }
                    }

                    // Optionally, redirect to another page after successful save
                    return RedirectToAction("Success");
                }
            }
            catch (Exception ex)
            {
                // Log exception for debugging
                System.Diagnostics.Debug.WriteLine("Error: " + ex.Message);
            }

            // If model state is not valid, return the view with the same model
            ViewBag.Message = "IMS - Integrated Management System";
            return View(model);
        }
  



        public ActionResult Contact()
        {
            List<demo> demoList = new List<demo>();
            string connectionString = ConfigurationManager.ConnectionStrings["MySqlConnection"].ConnectionString;
            string query = "SELECT Division, Department, Location, DocumentNo, DocumentName, Standards, Applicableclauses, Issueno, Revno, Preparer, Reviewer, Approver, SelectedDate, RevDate, DocDate,DocumentLevel FROM documents";

            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                using (MySqlCommand cmd = new MySqlCommand(query, conn))
                {
                    conn.Open();

                    using (MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            demo obj = new demo()
                            {
                                Division = reader["Division"].ToString(),
                                Department = reader["Department"].ToString(),
                                Location = reader["Location"].ToString(),
                                DocumentNo = reader["DocumentNo"].ToString(),
                                DocumentName = reader["DocumentName"].ToString(),
                                Standards = reader["Standards"].ToString(),
                                Applicableclauses = reader["Applicableclauses"].ToString(),
                                Issueno = reader["Issueno"].ToString(),
                                Revno = reader["Revno"].ToString(),
                                Preparer = reader["Preparer"].ToString(),
                                Reviewer = reader["Reviewer"].ToString(),
                                Approver = reader["Approver"].ToString(),
                                SelectedDate = (DateTime)(reader["SelectedDate"] != DBNull.Value ? Convert.ToDateTime(reader["SelectedDate"]) : (DateTime?)null),
                                Revdate = (DateTime)(reader["RevDate"] != DBNull.Value ? Convert.ToDateTime(reader["RevDate"]) : (DateTime?)null),
                                Docdate = (DateTime)(reader["DocDate"] != DBNull.Value ? Convert.ToDateTime(reader["DocDate"]) : (DateTime?)null),
                                DocumentLevel = reader["DocumentLevel"].ToString(),
                            };

                            demoList.Add(obj);
                        }
                    }
                }
            }

            return View(demoList);
        }

        // Action to download data as Excel
        public ActionResult DownloadExcel()
        {
            // Set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            List<demo> demoList = GetDemoData(); // Fetch the table data
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Document Data");

                // Add headers
                worksheet.Cells[1, 1].Value = "Division";
                worksheet.Cells[1, 2].Value = "Department";
                worksheet.Cells[1, 3].Value = "Location";
                worksheet.Cells[1, 4].Value = "Document No";
                worksheet.Cells[1, 5].Value = "Document Name";
                worksheet.Cells[1, 6].Value = "Standards";
                worksheet.Cells[1, 7].Value = "Applicable Clauses";
                worksheet.Cells[1, 8].Value = "Issue No";
                worksheet.Cells[1, 9].Value = "Rev No";
                worksheet.Cells[1, 10].Value = "Preparer";
                worksheet.Cells[1, 11].Value = "Reviewer";
                worksheet.Cells[1, 12].Value = "Approver";
                worksheet.Cells[1, 13].Value = "Issue Date";
                worksheet.Cells[1, 14].Value = "Rev Date";
                worksheet.Cells[1, 15].Value = "Document Date";
                worksheet.Cells[1, 16].Value = "Document Level";

                // Add data rows
                for (int i = 0; i < demoList.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = demoList[i].Division;
                    worksheet.Cells[i + 2, 2].Value = demoList[i].Department;
                    worksheet.Cells[i + 2, 3].Value = demoList[i].Location;
                    worksheet.Cells[i + 2, 4].Value = demoList[i].DocumentNo;
                    worksheet.Cells[i + 2, 5].Value = demoList[i].DocumentName;
                    worksheet.Cells[i + 2, 6].Value = demoList[i].Standards;
                    worksheet.Cells[i + 2, 7].Value = demoList[i].Applicableclauses;
                    worksheet.Cells[i + 2, 8].Value = demoList[i].Issueno;
                    worksheet.Cells[i + 2, 9].Value = demoList[i].Revno;
                    worksheet.Cells[i + 2, 10].Value = demoList[i].Preparer;
                    worksheet.Cells[i + 2, 11].Value = demoList[i].Reviewer;
                    worksheet.Cells[i + 2, 12].Value = demoList[i].Approver;
                    worksheet.Cells[i + 2, 13].Value = demoList[i].SelectedDate.ToShortDateString();
                    worksheet.Cells[i + 2, 14].Value = demoList[i].Revdate.ToShortDateString();
                    worksheet.Cells[i + 2, 15].Value = demoList[i].Docdate.ToShortDateString();
                    worksheet.Cells[i + 2, 16].Value = demoList[i].DocumentLevel;
                }

                var fileContents = package.GetAsByteArray();
                return File(fileContents, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DocumentData.xlsx");
            }
        }


        private List<demo> GetDemoData()
        {
            List<demo> demoList = new List<demo>();
            string connectionString = ConfigurationManager.ConnectionStrings["MySqlConnection"].ConnectionString;
            string query = "SELECT Division, Department, Location, DocumentNo, DocumentName, Standards, Applicableclauses, Issueno, Revno, Preparer, Reviewer, Approver, SelectedDate, RevDate, DocDate,DocumentLevel FROM documents";

            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                using (MySqlCommand cmd = new MySqlCommand(query, conn))
                {
                    conn.Open();

                    using (MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            demo obj = new demo()
                            {
                                Division = reader["Division"].ToString(),
                                Department = reader["Department"].ToString(),
                                Location = reader["Location"].ToString(),
                                DocumentNo = reader["DocumentNo"].ToString(),
                                DocumentName = reader["DocumentName"].ToString(),
                                Standards = reader["Standards"].ToString(),
                                Applicableclauses = reader["Applicableclauses"].ToString(),
                                Issueno = reader["Issueno"].ToString(),
                                Revno = reader["Revno"].ToString(),
                                Preparer = reader["Preparer"].ToString(),
                                Reviewer = reader["Reviewer"].ToString(),
                                Approver = reader["Approver"].ToString(),
                                SelectedDate = (DateTime)(reader["SelectedDate"] != DBNull.Value ? (DateTime?)reader["SelectedDate"] : null),
                                Revdate = (DateTime)(reader["RevDate"] != DBNull.Value ? (DateTime?)reader["RevDate"] : null),
                                Docdate = (DateTime)(reader["DocDate"] != DBNull.Value ? (DateTime?)reader["DocDate"] : null),
                                DocumentLevel = reader["DocumentLevel"].ToString(),
                            };

                            demoList.Add(obj);
                        }
                    }
                }
            }

            return demoList;
        }


        // POST: Home/Delete
        [HttpPost]
        public ActionResult Delete(int ID) // Change parameter type to string if Division is of type string
        {
            string connectionString = ConfigurationManager.ConnectionStrings["MySqlConnection"].ConnectionString;
            string query = "DELETE FROM documents WHERE ID = @ID"; // Using Division in the WHERE clause

            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                using (MySqlCommand cmd = new MySqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@ID", ID); // Bind the Division parameter
                    conn.Open();
                    cmd.ExecuteNonQuery(); // Execute the delete command
                }
            }

            return RedirectToAction("Index"); // Redirect to Index after deletion
        }
        ///
        public ActionResult Edit(int ID)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["MySqlConnection"].ConnectionString;
            demo document = null;

            string query = "SELECT * FROM documents WHERE ID = @ID";

            using (MySqlConnection con = new MySqlConnection(connectionString))
            {
                MySqlCommand cmd = new MySqlCommand(query, con);
                cmd.Parameters.AddWithValue("@ID", ID);

                con.Open();
                using (MySqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        document = new demo
                        {
                            ID = (int)reader["ID"],
                            Division = reader["Division"].ToString(),
                            Department = reader["Department"].ToString(),
                            Location = reader["Location"].ToString(),
                            DocumentNo = reader["DocumentNo"].ToString(),
                            DocumentName = reader["DocumentName"].ToString(),
                            Standards = reader["Standards"].ToString(),
                            Applicableclauses = reader["Applicableclauses"].ToString(),
                            Issueno = reader["Issueno"].ToString(),
                            Revno = reader["Revno"].ToString(),
                            Preparer = reader["Preparer"].ToString(),
                            Reviewer = reader["Reviewer"].ToString(),
                            Approver = reader["Approver"].ToString(),
                            SelectedDate = (DateTime)(reader["SelectedDate"] != DBNull.Value ? Convert.ToDateTime(reader["SelectedDate"]) : (DateTime?)null),
                            Revdate = (DateTime)(reader["Revdate"] != DBNull.Value ? Convert.ToDateTime(reader["Revdate"]) : (DateTime?)null),
                            Docdate = (DateTime)(reader["Docdate"] != DBNull.Value ? Convert.ToDateTime(reader["Docdate"]) : (DateTime?)null),
                            DocumentLevel = reader["DocumentLevel"].ToString()
                        };
                    }
                }
            }

            if (document != null)
            {
                return View(document); // Pass the document to the view for editing
            }
            return HttpNotFound();
        }
        [HttpPost]
        public ActionResult Save(demo model)
        {
            if (ModelState.IsValid)
            {
                string connectionString = ConfigurationManager.ConnectionStrings["MySqlConnection"].ConnectionString;

                using (MySqlConnection con = new MySqlConnection(connectionString))
                {
                    // Update the query to ensure you uniquely identify the document to update.
                    string query = @"UPDATE documents SET 
                    Division=@Division,
                    Department = @Department, 
                    Location = @Location, 
                    DocumentNo = @DocumentNo, 
                    DocumentName = @DocumentName, 
                    Standards = @Standards, 
                    Applicableclauses = @Applicableclauses,
                    Issueno = @Issueno, 
                    Revno = @Revno, 
                    Preparer = @Preparer, 
                    Reviewer = @Reviewer, 
                    Approver = @Approver, 
                    SelectedDate = @SelectedDate, 
                    Revdate = @Revdate, 
                    Docdate = @Docdate,
                    DocumentLevel=@DocumentLevel
                WHERE ID = @ID"; // or use an Id if available

                    MySqlCommand cmd = new MySqlCommand(query, con);
                    cmd.Parameters.AddWithValue("@ID", model.ID);
                    cmd.Parameters.AddWithValue("@Division", model.Division);
                    cmd.Parameters.AddWithValue("@Department", model.Department);
                    cmd.Parameters.AddWithValue("@Location", model.Location);
                    cmd.Parameters.AddWithValue("@DocumentNo", model.DocumentNo); // Make sure this is unique
                    cmd.Parameters.AddWithValue("@DocumentName", model.DocumentName);
                    cmd.Parameters.AddWithValue("@Standards", model.Standards);
                    cmd.Parameters.AddWithValue("@Applicableclauses", model.Applicableclauses);
                    cmd.Parameters.AddWithValue("@Issueno", model.Issueno);
                    cmd.Parameters.AddWithValue("@Revno", model.Revno);
                    cmd.Parameters.AddWithValue("@Preparer", model.Preparer);
                    cmd.Parameters.AddWithValue("@Reviewer", model.Reviewer);
                    cmd.Parameters.AddWithValue("@Approver", model.Approver);
                    cmd.Parameters.AddWithValue("@SelectedDate", model.SelectedDate);
                    cmd.Parameters.AddWithValue("@Revdate", model.Revdate);
                    cmd.Parameters.AddWithValue("@Docdate", model.Docdate);
                    cmd.Parameters.AddWithValue("@DocumentLevel", model.DocumentLevel);

                    try
                    {
                        con.Open();
                        int rowsAffected = cmd.ExecuteNonQuery(); // Execute the update command
                        if (rowsAffected > 0)
                        {
                            // Record updated successfully
                            return RedirectToAction("Index");
                        }
                        else
                        {
                            // Handle case where no record was updated
                            ModelState.AddModelError("", "Update failed, document not found.");
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log exception and return error message
                        ModelState.AddModelError("", "An error occurred: " + ex.Message);
                    }
                }
            }

            return View("Edit", model);
        }


        public ActionResult Success()
        {
            return View();  
        }

    }
}