using System;

namespace WecomMessageIntegration
{
    class Program
    {
        
        static void Main(string[] args)
        {
            var wecomHelper = new WecomHelper();
            var result = wecomHelper.SendMessage();
            Console.WriteLine(result);
        }
    }
}
