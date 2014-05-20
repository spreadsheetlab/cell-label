
namespace xlPrep
{
    class Program
    {
        static void Main(string[] args)
        {
            //Resolves to C:\Users\<username>\AppData\Roaming\xls
            //var filesPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "xls");
            var filesPath = "C:\\Euses";
            var outputPath = "C:\\EusesOutput";

            XlTransform t = new XlTransform();
            t.Transform(filesPath, outputPath);
        }
    }
}
