using System;
using dotnetconflib;
using dotnetconflib.EpplusSample.Sample01;

namespace dotnetconf
{
    class Program
    {
        static void Main(string[] args)
        {
           Sample01 demo = new Sample01();
           demo.RunSample1();
        }
    }
}
