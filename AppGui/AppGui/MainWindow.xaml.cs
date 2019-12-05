using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Xml.Linq;
using mmisharp;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using WindowsInput;
using WindowsInput.Native;

namespace AppGui
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private MmiCommunication mmiC;
        private IWebDriver driver;
        private InputSimulator sim;
        private List<string> tabs;

        private Int32 minWordLen;
        private Int32 minWordCount; 
        public MainWindow()
        {

            InitializeComponent();

            this.WindowState = WindowState.Minimized;
            this.ShowInTaskbar = false;

            minWordCount = 3;
            minWordLen = 4;

            driver = new ChromeDriver(".");
            driver.Manage().Window.Maximize();
            System.Threading.Thread.Sleep(3000);
            sim = new InputSimulator();
            tabs = driver.WindowHandles.ToList();

            mmiC = new MmiCommunication("localhost", 8000, "User1", "GUI");
            mmiC.Message += MmiC_Message;
            mmiC.Start();
            

        }

        private void MmiC_Message(object sender, MmiEventArgs e)
        {
            Console.WriteLine(e.Message);
            var doc = XDocument.Parse(e.Message);
            var com = doc.Descendants("command").FirstOrDefault().Value;
            dynamic json = JsonConvert.DeserializeObject(com);

            //We always receive 3 arguments -> commandID - commandName - confidence
            string commandID = json["recognized"][0];
            var commandName = json["recognized"][1];
            var confidence = json["recognized"][2];

            
            switch (commandID)
            {
                case "0":
                    //smth
                    break;
                case "1":
                    sim.Keyboard.KeyDown(VirtualKeyCode.CONTROL);
                    sim.Keyboard.KeyPress(VirtualKeyCode.VK_D);
                    sim.Keyboard.KeyUp(VirtualKeyCode.CONTROL);
                    System.Threading.Thread.Sleep(100);
                    sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
                    break;
                    
                case "2":
                   
                    driver.Manage().Window.Minimize();
                    break;
                    
                case "3":
                    driver.Manage().Window.Maximize();
                    break;
                    
                case "4":
                    sim.Keyboard.KeyDown(VirtualKeyCode.CONTROL);
                    sim.Keyboard.KeyPress(VirtualKeyCode.OEM_PLUS);
                    sim.Keyboard.KeyUp(VirtualKeyCode.CONTROL);
                    break;
                    
                case "5":
                    sim.Keyboard.KeyDown(VirtualKeyCode.CONTROL);
                    sim.Keyboard.KeyPress(VirtualKeyCode.OEM_MINUS);
                    sim.Keyboard.KeyUp(VirtualKeyCode.CONTROL);
                    break;
                case "6":
                    sim.Keyboard.KeyDown(VirtualKeyCode.CONTROL);
                    sim.Keyboard.KeyDown(VirtualKeyCode.SHIFT);
                    sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
                    sim.Keyboard.KeyUp(VirtualKeyCode.CONTROL);
                    sim.Keyboard.KeyUp(VirtualKeyCode.SHIFT);
                    tabs = driver.WindowHandles.ToList();
                    int index1 = tabs.IndexOf(driver.CurrentWindowHandle);
                    if (index1 == 0)
                    {
                        driver.SwitchTo().Window(driver.WindowHandles.Last());
                    }
                    else
                    {
                        driver.SwitchTo().Window(tabs[index1 - 1]);
                    }
                    break;
                case "7":
                    sim.Keyboard.KeyDown(VirtualKeyCode.CONTROL);
                    sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
                    sim.Keyboard.KeyUp(VirtualKeyCode.CONTROL);
                    //Console.WriteLine(tabs.ToString());
                    tabs = driver.WindowHandles.ToList();
                    int index2 = tabs.IndexOf(driver.CurrentWindowHandle);
                    if (index2 == tabs.Count() - 1)
                    {
                        driver.SwitchTo().Window(driver.WindowHandles.First());
                    }
                    else
                    {
                        driver.SwitchTo().Window(tabs[index2 + 1]);
                    }


                    break;
                case "8":

                    tabs = driver.WindowHandles.ToList();
                    if (tabs.Count() > 1)
                    {
                        driver.Close();
                        driver.SwitchTo().Window(driver.WindowHandles.First());
                    }

                    break;
                case "9":
                    string rec1 = "";
                    switch (rec1)
                    {
                        case "tab":
                            sim.Keyboard.KeyDown(VirtualKeyCode.CONTROL);
                            sim.Keyboard.KeyPress(VirtualKeyCode.VK_T);
                            sim.Keyboard.KeyUp(VirtualKeyCode.CONTROL);
                            driver.SwitchTo().Window(driver.WindowHandles.Last());
                            break;

                        case "historic":
                            sim.Keyboard.KeyDown(VirtualKeyCode.CONTROL);
                            sim.Keyboard.KeyPress(VirtualKeyCode.VK_H);
                            sim.Keyboard.KeyUp(VirtualKeyCode.CONTROL);
                            driver.SwitchTo().Window(driver.WindowHandles.Last());
                            break;
                        case "favourites":
                            sim.Keyboard.KeyDown(VirtualKeyCode.CONTROL);
                            sim.Keyboard.KeyDown(VirtualKeyCode.SHIFT);
                            sim.Keyboard.KeyPress(VirtualKeyCode.VK_O);
                            sim.Keyboard.KeyUp(VirtualKeyCode.CONTROL);
                            sim.Keyboard.KeyUp(VirtualKeyCode.SHIFT);
                            break;
                    }
                    break;
            }


        }
    }
}