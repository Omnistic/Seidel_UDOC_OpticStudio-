using System;
using System.IO;
using ZOSAPI;
using ZOSAPI.Analysis;
using ZOSAPI.Analysis.Data;
using ZOSAPI.Analysis.Settings.Aberrations;


namespace CSharpUserOperandApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            // Find the installed version of OpticStudio
            bool isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize();
            // Note -- uncomment the following line to use a custom initialization path
            //bool isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize(@"C:\Program Files\OpticStudio\");
            if (isInitialized)
            {
                LogInfo("Found OpticStudio at: " + ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory());
            }
            else
            {
                HandleError("Failed to locate OpticStudio!");
                return;
            }

            BeginUserOperand();
        }

        static void BeginUserOperand()
        {
            // Create the initial connection class
            ZOSAPI_Connection TheConnection = new ZOSAPI_Connection();

            // Attempt to connect to the existing OpticStudio instance
            IZOSAPI_Application TheApplication = null;
            try
            {
                TheApplication = TheConnection.ConnectToApplication(); // this will throw an exception if not launched from OpticStudio
            }
            catch (Exception ex)
            {
                HandleError(ex.Message);
                return;
            }
            if (TheApplication == null)
            {
                HandleError("An unknown connection error occurred!");
                return;
            }

            // Check the connection status
            if (!TheApplication.IsValidLicenseForAPI)
            {
                HandleError("Failed to connect to OpticStudio: " + TheApplication.LicenseStatus);
                return;
            }
            if (TheApplication.Mode != ZOSAPI_Mode.Operand)
            {
                HandleError("User plugin was started in the wrong mode: expected Operand, found " + TheApplication.Mode.ToString());
                return;
            }

            // Read the operand arguments

            // Hx column
            int wave = (int) TheApplication.OperandArgument1;

            // Hy column
            int surface = (int) TheApplication.OperandArgument2;

            // Initialize the output array
            int maxResultLength = TheApplication.OperandResults.Length;
            double[] operandResults = new double[maxResultLength];

            IOpticalSystem TheSystem = TheApplication.PrimarySystem;
            // Add your custom code here...

            // Basic error trapping for arguments

            // Hx: wave number, by default uses the wavelength 1, as far as I know there's no primary in Seidel Coefficients
            if (wave < 1 || wave > TheSystem.SystemData.Wavelengths.NumberOfWavelengths) wave = 1;

            // Hy: surface number, by default uses 0 meaning the sum of coefficients for all surfaces
            if (surface < 0 || surface > TheSystem.LDE.NumberOfSurfaces - 1) surface = 0;

            // Create a Seidel Coefficients analysis
            IA_ seidel_coefficients_analysis = TheSystem.Analyses.New_SeidelCoefficients();

            // Adjust the wavelength according to the Hx column
            IAS_SeidelCoefficients seidel_coefficients_analysis_settings = seidel_coefficients_analysis.GetSettings() as IAS_SeidelCoefficients;
            seidel_coefficients_analysis_settings.Wavelength.SetWavelengthNumber(wave);

            // Run the Seidel Coefficients analysis
            seidel_coefficients_analysis.ApplyAndWaitForCompletion();

            // Retrieve the results of the Seidel Coefficients analysis
            IAR_ seidel_coefficients_analysis_results = seidel_coefficients_analysis.GetResults();

            // Save the results of the Seidel Coefficients analysis to a text file
            string temporary_filename = @"seidel_coefficients_temp.txt";
            string full_path = Path.Combine(TheApplication.SamplesDir, temporary_filename);
            seidel_coefficients_analysis_results.GetTextFile(full_path);

            // Start parsing the text file...
            string new_line = "";
            StreamReader sr = new StreamReader(full_path);

            // How many more lines aside from the header are to be ignored?
            // If surface == 0 we ignore all the lines until we reach TOT
            int ignore_offset = surface - 1;
            if (surface == 0) ignore_offset = TheSystem.LDE.NumberOfSurfaces - 1;

            // Ignore header lines and lines from previous surfaces
            for (int ii = 0; ii < 18 + ignore_offset; ii++)
            {
                new_line = sr.ReadLine();
            }

            // Parse the relevant surface line
            string[] surface_coefficients_string = sr.ReadLine().Split('\t');

            // Save the CLA as Data 0
            operandResults[0] = Convert.ToDouble(surface_coefficients_string[6]);

            // Save the CTR as Data 1
            operandResults[1] = Convert.ToDouble(surface_coefficients_string[7]);

            // Close streamreader
            sr.Close();

            // Delete temporary text file
            File.Delete(full_path);

            // Clean up
            FinishUserOperand(TheApplication, operandResults);
        }

        static void FinishUserOperand(IZOSAPI_Application TheApplication, double[] resultData)
        {
            // Note - OpticStudio will wait for the operand to complete until this application exits 
            if (TheApplication != null)
            {
                TheApplication.OperandResults.WriteData(resultData.Length, resultData);
            }
        }

        static void LogInfo(string message)
        {
            // TODO - add custom logging
            Console.WriteLine(message);
        }

        static void HandleError(string errorMessage)
        {
            // TODO - add custom error handling
            throw new Exception(errorMessage);
        }

    }
}
