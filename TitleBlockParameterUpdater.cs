using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Linq;

namespace TitleblockKeyplanUpdater
{
    [Transaction(TransactionMode.Manual)]
    public class TitleBlockParameterUpdater : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                // Get the active document
                Document doc = commandData.Application.ActiveUIDocument.Document;

                // Prompt user to select a title block family instance
                FamilyInstance selectedTitleBlock = GetTitleBlockInstance(commandData);
                if (selectedTitleBlock == null)
                    return Result.Cancelled;

                // Prompt user for input text
                string userInput = GetUserInput("Enter text to match sheet names:");
                if (string.IsNullOrEmpty(userInput))
                    return Result.Cancelled;

                // Prompt user to select a Yes/No parameter
                string selectedParameter = GetYesNoParameter(doc, selectedTitleBlock);
                if (string.IsNullOrEmpty(selectedParameter))
                    return Result.Cancelled;

                // Get all sheets in the document
                FilteredElementCollector sheetCollector = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet));

                // Iterate through each sheet and update the title block parameter
                foreach (ViewSheet sheet in sheetCollector)
                {
                    if (sheet.Name.Contains(userInput))
                    {
                        // Update the title block parameter for the matching sheet
                        using (Transaction transaction = new Transaction(doc, "Update Title Block Parameter"))
                        {
                            transaction.Start();

                            // Find the title block instances on the sheet
                            IEnumerable<FamilyInstance> titleBlocksOnSheet = new FilteredElementCollector(doc, sheet.Id)
                                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                                .OfClass(typeof(FamilyInstance))
                                .Cast<FamilyInstance>();

                            // Assuming the title block is the first one on the sheet
                            FamilyInstance titleBlock = titleBlocksOnSheet.FirstOrDefault();


                            if (titleBlock != null)
                            {
                                // Update the selected parameter
                                Parameter parameter = titleBlock.LookupParameter(selectedParameter);
                                if (parameter != null && parameter.StorageType == StorageType.Integer)
                                {
                                    // Assuming it's a Yes/No parameter
                                    parameter.Set(1); // 1 for true (Yes), 0 for false (No)
                                }
                            }

                            transaction.Commit();
                        }
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private string GetUserInput(string prompt)
        {
            // Create a WinForms form for user input
            System.Windows.Forms.Form userInputForm = new System.Windows.Forms.Form();
            System.Windows.Forms.Label label = new System.Windows.Forms.Label();
            System.Windows.Forms.TextBox textBox = new System.Windows.Forms.TextBox();
            System.Windows.Forms.Button okButton = new System.Windows.Forms.Button();

            // Set the label and button properties
            label.Text = prompt;
            label.Location = new System.Drawing.Point(10, 10);
            label.Size = new System.Drawing.Size(300, 20);

            textBox.Location = new System.Drawing.Point(10, 40);
            textBox.Size = new System.Drawing.Size(300, 20);

            okButton.Text = "OK";
            okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            okButton.Location = new System.Drawing.Point(10, 70);

            // Add controls to the form
            userInputForm.Controls.Add(label);
            userInputForm.Controls.Add(textBox);
            userInputForm.Controls.Add(okButton);

            // Event handler for Enter key
            userInputForm.AcceptButton = okButton;

            string result = null;

            // Show the form and get user input
            if (userInputForm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                result = textBox.Text;
            }

            return result;
        }

        private FamilyInstance GetTitleBlockInstance(ExternalCommandData commandData)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Selection sel = uiDoc.Selection;

            // Prompt user to select a title block family instance
            Reference pickedRef = sel.PickObject(ObjectType.Element, new TitleBlockSelectionFilter(),
                "Select a Title Block");

            if (pickedRef == null || pickedRef.ElementId == ElementId.InvalidElementId)
                return null; // User canceled the operation or selected an invalid element

            return uiDoc.Document.GetElement(pickedRef) as FamilyInstance;
        }

        private string GetYesNoParameter(Document doc, FamilyInstance titleBlock)
        {
            // Get all parameters of the selected title block
            List<string> yesNoParameters = new List<string>();
            foreach (Parameter parameter in titleBlock.Parameters)
            {
                // Check if the parameter is of Yes/No type
                if (parameter.StorageType == StorageType.Integer)
                {
                    yesNoParameters.Add(parameter.Definition.Name);
                }
            }

            // Create a WinForms form for parameter selection
            System.Windows.Forms.Form parameterSelectionForm = new System.Windows.Forms.Form();
            System.Windows.Forms.Label label = new System.Windows.Forms.Label();
            System.Windows.Forms.ComboBox comboBox = new System.Windows.Forms.ComboBox();
            System.Windows.Forms.Button okButton = new System.Windows.Forms.Button();

            // Set the label and button properties
            label.Text = "Choose a Yes/No parameter";
            label.Location = new System.Drawing.Point(10, 10);
            label.Size = new System.Drawing.Size(300, 20);

            comboBox.Location = new System.Drawing.Point(10, 40);
            comboBox.Size = new System.Drawing.Size(300, 20);

            okButton.Text = "OK";
            okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            okButton.Location = new System.Drawing.Point(10, 70);

            // Add controls to the form
            parameterSelectionForm.Controls.Add(label);
            parameterSelectionForm.Controls.Add(comboBox);
            parameterSelectionForm.Controls.Add(okButton);

            // Event handler for Enter key
            parameterSelectionForm.AcceptButton = okButton;

            // Add Yes/No parameters as options to the ComboBox
            comboBox.Items.AddRange(yesNoParameters.ToArray());

            string result = null;

            // Show the form and get user input
            if (parameterSelectionForm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                result = comboBox.SelectedItem as string;
            }

            return result;
        }
    }

    // Custom selection filter to allow only FamilyInstances with a Title Block
    public class TitleBlockSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is FamilyInstance familyInstance &&
                   familyInstance.Symbol.Family.FamilyCategory.Id.IntegerValue ==
                   (int)BuiltInCategory.OST_TitleBlocks;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}
