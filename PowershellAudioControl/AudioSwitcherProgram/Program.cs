namespace AudioSwitcherProgram
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string result = PowerShellController.SendCommand("get-host | select-object version");
            Console.WriteLine(result);
        }
    }
}