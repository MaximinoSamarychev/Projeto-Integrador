using Microsoft.Win32;
using Siemens.Engineering;
using Siemens.Engineering.Cax;
using Siemens.Engineering.Compiler;
using Siemens.Engineering.Compare;
using Siemens.Engineering.Download;
using Siemens.Engineering.Hmi;
using Siemens.Engineering.HmiUnified;
using Siemens.Engineering.Hmi.Cycle;
using Siemens.Engineering.Hmi.Communication;
using Siemens.Engineering.Hmi.Globalization;
using Siemens.Engineering.Hmi.RuntimeScripting;
using Siemens.Engineering.Hmi.Screen;
using Siemens.Engineering.Hmi.Tag;
using Siemens.Engineering.Hmi.TextGraphicList;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Extensions;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.HW.Utilities;
using Siemens.Engineering.Library;
using Siemens.Engineering.Library.MasterCopies;
using Siemens.Engineering.Library.Types;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW.ExternalSources;
using Siemens.Engineering.SW.Tags;
using Siemens.Engineering.SW.TechnologicalObjects;
using Siemens.Engineering.SW.TechnologicalObjects.Motion;
using Siemens.Engineering.SW.Types;
using Siemens.Engineering.Upload;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using Siemens.Engineering.HmiUnified.UI.Controls;
using Siemens.Engineering.HmiUnified.UI.Events;
using Siemens.Engineering.HmiUnified.UI.Parts;
using System.Runtime.CompilerServices;
using Siemens.Engineering.Hmi.Faceplate;
using Siemens.Engineering.HmiUnified.Common;
using System.Security.Policy;
using Siemens.Engineering.HmiUnified.UI.ScreenGroup;
using Siemens.Engineering.HmiUnified.UI.Screens;
using Siemens.Engineering.SW.Blocks.Interface;
using System.Linq.Expressions;
using Siemens.Engineering.SW.Blocks.Exceptions;
using System.Xml;
using static System.Net.WebRequestMethods;









namespace OpenessTIA
{
    public class TIA_V18
    {


        public static TiaPortal instTIA;
        public static Project projectTIA;
        
        public static Device plcDevice;
        public static Device hmiDevice;
        public static Device hmiUnifiedDevice;
        public static PlcSoftware plcSoftware;
        public static HmiSoftware hmiSoftware;
        public static HmiTarget hmiTarget;
        public static ProjectLibrary projectLibrary;
        public static DeviceItem plcDeviceItem;
        public static DeviceItem hmiDeviceItem;
        public static int numeroDBsCylinder = 0;
        public static int numeroDBs = 0;
        public static int numeroFCs = 0;
        
        

        public static HmiFaceplateInterfaceComposition hmiFaceplateInterfaceComp;
        
        public static HmiFaceplateContainer hmiFaceplateContainer;
        public static HmiFaceplateInterface hmiFaceplate;

        public static LibraryTypeFolder libraryTypeFolder;
        public static FaceplateLibraryType faceplateFolder;


        public static Node no;
        public static Subnet subnet;
        public static UserGlobalLibrary globalLibrary;
        public static HmiScreen hmiScreen;






        #region Gets e Sets

        public int getNumeroDBsCylinder()
        {
            return numeroDBsCylinder;
        }
        #endregion


        #region Abertura do TIA Portal e do Projeto e salvar projeto
        //Creates new TIA Portal instance with/without user interface
        public void createTiaInstance(bool guiTIA)
        {
            //whitelist entry
            SetWhitelist(System.Diagnostics.Process.GetCurrentProcess().ProcessName, System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);

            if (guiTIA)
            {
                instTIA = new TiaPortal(TiaPortalMode.WithUserInterface);
                if (instTIA == null) Console.WriteLine("Instance null inicio");
                
            }
            else
            {
                instTIA = new TiaPortal(TiaPortalMode.WithoutUserInterface);
                if (instTIA == null) Console.WriteLine("Instance null inicio");
            }
        }


        //Creates or Opens a project in TIA Portal

        public void createOpenTiaProject(string projectPath, string projectName, bool createOpen)
        {
           
                
            if (instTIA == null) Console.WriteLine("Instance null");

            //Create the project with specified directory and project name and opens it automatically
            if (createOpen == false)
            {
                //Specify the directory where the project will be created
                DirectoryInfo targetDirectory = new DirectoryInfo(projectPath);
                Console.WriteLine(targetDirectory);
                projectTIA = instTIA.Projects.Create(targetDirectory, projectName);
                Console.WriteLine("Project created");
            }
            else
            {


                //Open Project with specified path
                FileInfo targetFile = new FileInfo(projectPath + "\\" + projectName + "\\" + projectName + ".ap18");
                Console.WriteLine(targetFile);
                projectTIA = instTIA.Projects.Open(targetFile);
                Console.WriteLine("Project oppened");
            }

        }

        public void openProjectView()
        {


            projectTIA.ShowHwEditor(View.Network);
        }

        static void SetWhitelist(string ApplicationName, string ApplicationStartupPath)
        {

            RegistryKey key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            RegistryKey software = null;
            try
            {
                software = key.OpenSubKey(@"SOFTWARE\Siemens\Automation\Openness")
                    .OpenSubKey("18.0")
                    .OpenSubKey("Whitelist")
                    .OpenSubKey(ApplicationName + ".exe")
                    .OpenSubKey("Entry", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryRights.FullControl);
            }
            catch (Exception)
            {

                //Eintrag in der Whitelist ist nicht vorhanden
                //Entry in whitelist is not available
                software = key.CreateSubKey(@"SOFTWARE\Siemens\Automation\Openness")
                    .CreateSubKey("18.0")
                    .CreateSubKey("Whitelist")
                    .CreateSubKey(ApplicationName + ".exe")
                    .CreateSubKey("Entry", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryOptions.None);
            }


            string lastWriteTimeUtcFormatted = String.Empty;
            DateTime lastWriteTimeUtc;
            HashAlgorithm hashAlgorithm = SHA256.Create();
            FileStream stream = System.IO.File.OpenRead(ApplicationStartupPath);
            byte[] hash = hashAlgorithm.ComputeHash(stream);
            // this is how the hash should appear in the .reg file
            string convertedHash = Convert.ToBase64String(hash);
            software.SetValue("FileHash", convertedHash);
            lastWriteTimeUtc = new FileInfo(ApplicationStartupPath).LastWriteTimeUtc;
            // this is how the last write time should be formatted
            lastWriteTimeUtcFormatted = lastWriteTimeUtc.ToString(@"yyyy/MM/dd HH:mm:ss.fff", CultureInfo.InvariantCulture);
            software.SetValue("DateModified", lastWriteTimeUtcFormatted);
            software.SetValue("Path", ApplicationStartupPath);
        }

        public void saveProject()
        {
            projectTIA.Save();
        }
        #endregion Abertura do TIA Portal e do Projeto

        #region Criar e Encontrar Hmi's e Plc's no projeto


        // Creates PLC with a given name
        public void createDevicePlc(string plcName = "PLC")
        {


            string plcVersion = "V3.0"; 
            string plcArticle = "6ES7 512-1SM03-0AB0";
            string plcIdentifier = "OrderNumber:" + plcArticle + "/" + plcVersion;
            string plcStation = "station" + plcName;

            //Creates new PLC with specified Version and Acrticle in TIA Project
            plcDevice = projectTIA.Devices.CreateWithItem(plcIdentifier, plcName, plcStation);

            //Obtem o Device Item
            plcDeviceItem = plcDevice.DeviceItems.First(Device => Device.Name.Equals(plcName));

            plcDevice.SetAttribute("Name", "PLC");
        }

        //Creates HMI with a given name
        public void createDeviceHMI(bool unified , string hmiName = "HMI")
        {
           
            string hmiVersion = "";
            string hmiArcticle = "";
            string hmiIdentifier = "";
            string hmiStation = "";
            

            if (unified == false)
            {
                hmiVersion = "17.0.0.0";
                hmiArcticle = "6AV2 124-0MC01-0AX0";
                hmiIdentifier = "OrderNumber:" + hmiArcticle + "/" + hmiVersion;
                hmiStation = null;
                hmiDevice = projectTIA.Devices.CreateWithItem(hmiIdentifier, hmiName, hmiStation);
            }
            else
            {
                hmiVersion = "18.0.0.0";
                hmiArcticle = "6AV2 128-3MB06-0AXx";
                hmiIdentifier = "OrderNumber:" + hmiArcticle + "/" + hmiVersion;
                hmiStation = null;
                hmiName = "UnifiedHMI";
                hmiUnifiedDevice = projectTIA.Devices.CreateWithItem(hmiIdentifier, hmiName, hmiStation);

            }
            
            

            
        }

        //Encontra os devices do projeto (Retorna 1 se não encontrar PLC, 2 se não encontrar HMI e 3 se não encontrar nem PLC e nem HMI e Retorna 0 se encontrar os dois)
        public int findDevices(string plcName, string hmiName)
        {

            int erro = 0;
            int count = 0;
            int numDevices = projectTIA.Devices.Count;

            if (numDevices == 0)
            {
                erro = 3;
            }else 
             { 


                for (count = 0; count < numDevices; count++)
                {



                    Device auxDevice = projectTIA.Devices[count];

                    if (auxDevice.Name == plcName)
                    {
                        plcDevice = auxDevice;

                        Console.WriteLine(plcDevice.Name + " Found");


                        foreach (DeviceItem deviceItem in plcDevice.DeviceItems)
                        {
                            SoftwareContainer softwareContainer = ((IEngineeringServiceProvider)deviceItem).GetService<SoftwareContainer>();

                            if (softwareContainer != null)
                            {
                                plcDeviceItem = plcDevice.DeviceItems.First(Device => Device.Name.Equals(plcName));
                                if (plcDeviceItem == null) Console.WriteLine("Plc Device Item is null");

                            }
                        }

                    }

                    else if (auxDevice.Name == hmiName)
                    {
                        hmiDevice = auxDevice;

                        Console.WriteLine(hmiDevice.Name + " Found");
                        foreach (DeviceItem deviceItem in hmiDevice.DeviceItems)
                        {
                            SoftwareContainer softwareContainer = ((IEngineeringServiceProvider)deviceItem).GetService<SoftwareContainer>();

                            if (softwareContainer != null)
                            {
                                hmiDeviceItem = hmiDevice.DeviceItems.First(Device => Device.Name.Equals(hmiName));
                                if (hmiDeviceItem == null) Console.WriteLine("Hmi Device Item is null");
                            }
                        }
                    }

                }
            }

                if (plcDevice == null && hmiDevice != null)
                {
                    Console.WriteLine("PLC not found");
                    erro = 1;
                } else if (plcDevice != null && hmiDevice == null)
                {
                    Console.WriteLine("HMI not found");
                    erro = 2;
                } else if (plcDevice == null && hmiDevice == null)
                {
                    Console.WriteLine("PLC and HMI not Found");
                    erro = 3;
                }

                

            
            
            return erro;
        }


        #endregion Criar e Encontrar Hmi's e Plc's no projeto


        #region Obtenção do software dos devices
        public void getPlcSoftware()
        {

            foreach (DeviceItem deviceItem in plcDevice.DeviceItems)
            {
                SoftwareContainer softwareContainer = ((IEngineeringServiceProvider)deviceItem).GetService<SoftwareContainer>();


                if (softwareContainer != null)
                {
                    plcSoftware = (PlcSoftware)softwareContainer.Software;




                }

            }


        }

        //Obtem o hmi Target se a HMi não for Unified
        public void getHmiTarget()
        {


            foreach (DeviceItem deviceItem in hmiDevice.DeviceItems)
            {
                SoftwareContainer softwareContainer = ((IEngineeringServiceProvider)deviceItem).GetService<SoftwareContainer>();

                if (softwareContainer != null)
                {
                    hmiTarget = (HmiTarget)softwareContainer.Software;


                }

            }



        }

        //Obtem HmiSoftware se a HMI for Unified
        public void getHmiSoftware()
        {
            var deviceItems = hmiDevice.DeviceItems;
            if (deviceItems != null)
            {
                foreach (DeviceItem deviceItem in deviceItems)
                {
                    Console.WriteLine(deviceItem.Name);
                    SoftwareContainer softwareContainer = deviceItem.GetService<SoftwareContainer>();
                    hmiSoftware = softwareContainer?.Software as HmiSoftware;

                    if (hmiSoftware == null)
                    {
                        Console.WriteLine("hmiSoftware is null");
                    }
                    else
                    {
                        Console.WriteLine(hmiSoftware.Name);
                        Console.WriteLine("hmiSoftware is not null");
                    }
                }
            }




            Console.WriteLine(hmiSoftware.Name);
        }

        #endregion Obtenção do software dos devices

        #region Conexão e atribuição de IP's



        //Dá um IP ao PLC, cria e concecta à subnet com nome especificado | Esta função deve ser executada antes da giveHmiIpAddress()
        public void givePlcIPAddress(string ipAddress, string subnetName = "PN/IE_1")
        {

            

            DeviceItem plcProfinet = plcDeviceItem.DeviceItems.First(DeviceItem => DeviceItem.Name.Equals("PROFINET interface_1") );
            
            NetworkInterface plcNetworkInterface = ((IEngineeringServiceProvider)plcProfinet).GetService<NetworkInterface>();

            if(plcNetworkInterface != null)
            {
                foreach(Node node in plcNetworkInterface.Nodes)
                {
                    

                    if (node != null)
                    {
                        foreach(EngineeringAttributeInfo nodeInfo in node.GetAttributeInfos())
                        {
                            

                            if(nodeInfo != null && nodeInfo.Name == "Address" && ipAddress != null)
                            {
                                node.SetAttribute("Address", ipAddress);
                                Console.WriteLine("Given " + ipAddress + " to " + plcDevice.Name);
                                if (node.ConnectedSubnet == null)
                                {
                                    subnet = node.CreateAndConnectToSubnet(subnetName);
                                    Console.WriteLine("Created and Connected to subnet " + subnetName);
                                }
                                else
                                {
                                    Console.WriteLine("PLC already connected to subnet " + node.ConnectedSubnet.Name);
                                }
                                
                            }
                        }
                    }
                }

                
            }


        }

        //Dá um IP à HMI e conecta à subnet do projeto
        public void giveHmiIPAddress(string ipAddress)
        {
            DeviceItem hmiDeviceItemForIp = hmiDevice.DeviceItems.First(Device => Device.Name.Equals("HMI.IE_CP_1"));

            DeviceItem hmiProfinet = hmiDeviceItemForIp.DeviceItems.First(DeviceItem => DeviceItem.Name.Equals("PROFINET Interface_1"));



            NetworkInterface hmiNetworkInterface = ((IEngineeringServiceProvider)hmiProfinet).GetService<NetworkInterface>();
            if (hmiNetworkInterface == null) Console.WriteLine("hmi Network Interface is null");
            
            if (hmiNetworkInterface != null)
            {
                
                foreach (Node node in hmiNetworkInterface.Nodes)
                {

                    if (node != null)
                    {
                        
                        foreach (EngineeringAttributeInfo nodeInfo in node.GetAttributeInfos())
                        {
          

                            if (nodeInfo != null && nodeInfo.Name == "Address" && ipAddress != null)
                            {
                               
                                node.SetAttribute("Address", ipAddress);
                                Console.WriteLine("Given " + ipAddress + " to " + hmiDevice.Name);
                                if (node.ConnectedSubnet == null)
                                {
                                    node.ConnectToSubnet(subnet);
                                    Console.WriteLine("Connected to subnet " + subnet.Name);
                                }
                                else
                                {
                                    Console.WriteLine("HMI already connected to subnet " + node.ConnectedSubnet.Name);
                                }


                            }
                        }
                    }
                }


            }
            

        }

        //usa as funcções givePlcIPAddress e giveHmiIPAddress e conecta-as à subnet 
        public void connectDevices(string plcIp = "192.168.192.1", string hmiIp = "192.168.192.2", string subnetName = "PN/IE_1")
        {
            givePlcIPAddress(plcIp, subnetName);

            giveHmiIPAddress(hmiIp);

        }

        #endregion Conexão e atribuição de IP's



        #region Funções base para criação de pastas importar Global Library e Importar objetos da Global Library
        public void countDataBlocks()
        {
            numeroDBs = 0;
            numeroDBsCylinder = 0;
            foreach (PlcBlock block in plcSoftware.BlockGroup.Blocks)
            {
                if(block is DataBlock)
                {
                    numeroDBs++;

                    if (block.Name.Substring(0, "Cilindro".Length) == "Cilindro")
                    {

                        numeroDBsCylinder++;
                    }
                }
            }
            foreach (PlcBlock block in plcSoftware.BlockGroup.Groups.Find("DataBlocks").Blocks)
            {
                if (block is DataBlock)
                {
                    
                    numeroDBs++;
                    
                    if (block.Name.Substring(0,"Cilindro".Length) == "Cilindro" )
                    {
                        
                        numeroDBsCylinder++;
                    }
                }
            }

            Console.WriteLine("Numero de DataBlocks: " + numeroDBs);
            Console.WriteLine("Numero de DataBlocks Cilindro: " + numeroDBsCylinder);
        }
        //Itens do PLC

        public void countFCs()
        {
            numeroFCs = 0;

            foreach (PlcBlock block in plcSoftware.BlockGroup.Groups.Find("FCs").Blocks)
            {
                if(block is FC)
                {
                    numeroFCs++;
                }
                
            }

            foreach (PlcBlock block in plcSoftware.BlockGroup.Groups.Find("FCs").Blocks)
            {
                if(block is FC)
                {
                    numeroFCs++;
                }
                
            }

            Console.WriteLine("Numero de FC's: "+ numeroFCs);
        }

        //Importa a Global Library com Faceplates UDts e Fb's da Controlar 
        public void importGlobalLibrary(string libraryAddress)
        {
            FileInfo info = new FileInfo(libraryAddress);
            globalLibrary = instTIA.GlobalLibraries.Open(info, OpenMode.ReadWrite);

            Console.WriteLine("Global Library imported");
        }

        public void createPlcFolders()
        {
            var plcFolder = plcSoftware.BlockGroup.Groups;
            int numFolders = plcFolder.Count;
            bool[] existeFolder = { false, false, false };
            if (numFolders == 0)
            {
                plcFolder.Create("FBs");
                plcFolder.Create("DataBlocks");
                plcFolder.Create("FCs");
                Console.WriteLine("Folders Created");
            }
            else
            {
                for(int i = 0; i < numFolders; i++)
                {
                    if (plcFolder[i].Name == "FBs")
                    {
                        existeFolder[0] = true;
                    }
                    else if (plcFolder[i].Name == "DataBlocks")
                    {
                        existeFolder[1] = true;
                    }
                    else if (plcFolder[i].Name == "FCs")
                    {
                        existeFolder[2] = true;
                    }
                }

                if (existeFolder[0] == false)
                {
                    plcFolder.Create("FBs");
                    Console.WriteLine("Folder FBs Created");
                } 
                if(existeFolder[1] == false)
                {
                    plcFolder.Create("DataBlocks");
                    Console.WriteLine("Folder DataBlocks Created");
                }
                if(existeFolder[2] == false)
                {
                    plcFolder.Create("FCs");
                    Console.WriteLine("Folder FCs Created");
                }
            }
        }

        //Cria UDT's a Partir das Master Copies Cria também pasta "UDTs" caso não exista    
        public void getUdtFromLibrary()
        {

            int existeFolder = 0;
            var plcFolder = plcSoftware.TypeGroup.Groups;
            int numCopies = globalLibrary.MasterCopyFolder.Folders.Find("UDTs").MasterCopies.Count;

            int numPlcFolders = plcFolder.Count;


            if (numPlcFolders > 0)
            {


                for (int i = 0; i < numPlcFolders; i++)
                {
                    if (plcFolder[i].Name == "UDTs")
                    {
                        existeFolder = 1;
                    }
                }
            }
            if (numCopies != 0) { 
                if (existeFolder == 0)
                {
                    plcFolder.Create("UDTs");
                    Console.WriteLine("--->Group UDTs Created");
                }

                foreach (PlcTypeUserGroup group in plcFolder)
                {
                    
                    if (group.Name == "UDTs")
                    {


                        var udtFolder = group.Groups;



                        PlcTypeComposition typeComposition = group.Types;


                        for (int i = 0; i < numCopies; i++)
                        {
                            MasterCopy masterCopySource = globalLibrary.MasterCopyFolder.Folders.Find("UDTs").MasterCopies[i];




                            bool existe = false;
                            foreach (PlcType type in typeComposition)
                            {
                                if (type.Name == masterCopySource.Name)
                                {
                                    existe = true;
                                }



                            }

                            if (existe == false)
                            {
                                group.Types.CreateFrom(masterCopySource);
                                Console.WriteLine("UDT " + masterCopySource.Name + " Created in folder " + group.Name);
                            }
                        }
                    }

                }

            }
            else
            {
                Console.WriteLine("There are no User-Defined Types in Master Copies");
            }
        }

        //Cria FB's a Partir das Master Copies Cria também pasta "FBs" caso não exista 
        public void getFbFromLibrary()
        {
            int existeFolder = 0;
            int numCopies = globalLibrary.MasterCopyFolder.Folders.Find("FBs").MasterCopies.Count;
            var plcFolder = plcSoftware.BlockGroup.Groups;
            int numFbFolders = plcFolder.Count;


            if (numCopies != 0)
            {

                if (numFbFolders > 0)
                {
                    for (int i = 0; i < numFbFolders; i++)
                    {
                        if (plcFolder[i].Name == "FBs")
                        {
                            existeFolder = 1;
                        }
                    }
                }

                if (existeFolder == 0)
                {
                    plcFolder.Create("FBs");
                    Console.WriteLine("--->Folder FBs Created");
                }

                foreach (PlcBlockUserGroup group in plcFolder)
                {
                    
                    if (group.Name == "FBs")
                    {
                        var fbFolder = group.Groups;

                        

                        PlcBlockComposition blockComposition = group.Blocks;

                        

                        for (int i = 0; i < numCopies; i++)
                        {
                            MasterCopy masterCopySource = globalLibrary.MasterCopyFolder.Folders.Find("FBs").MasterCopies[i];




                            bool existe = false;
                            foreach (PlcBlock block in blockComposition)
                            {
                                if (block.Name == masterCopySource.Name)
                                {
                                    existe = true;
                                }



                            }

                            if (existe == false)
                            {
                                group.Blocks.CreateFrom(masterCopySource);
                                Console.WriteLine("Block " + masterCopySource.Name + " Created in folder " + group.Name);
                            }
                        }
                    }


                }
            }
            else
            {
                Console.WriteLine("There are no Function Block Master Copies");
            }

        }

        //Importa DataBlocks da Global Library. Para Datablocks de Objetos funciona. Implementar XML depois
        public int getDataBlockFromLibrary(string dbName)
        {
            int existeFolder = 0;
            int numCopies = globalLibrary.MasterCopyFolder.Folders.Find("DataBlocks").MasterCopies.Count;
            var plcFolder = plcSoftware.BlockGroup.Groups;
            int numBlockFolders = plcFolder.Count;
            bool existeCopia = false;

            for(int i = 0; i < numCopies; i++)
            {
                if(globalLibrary.MasterCopyFolder.Folders.Find("DataBlocks").MasterCopies[i].Name == dbName)
                {
                    existeCopia = true;
                }
            }

            if(existeCopia == false)
            {
                Console.WriteLine("The Copy " + dbName + " doesn't exists");
                return -1;
            }

            if (numCopies != 0)
            {

                if (numBlockFolders > 0)
                {
                    for (int i = 0; i < numBlockFolders; i++)
                    {
                        if (plcFolder[i].Name == "DataBlocks")
                        {
                            existeFolder = 1;
                        }
                    }
                }

                if (existeFolder == 0)
                {
                    plcFolder.Create("DataBlocks");
                    Console.WriteLine("--->Folder DataBlocks Created");
                }

                foreach (PlcBlockUserGroup group in plcFolder)
                {
                    
                    if (group.Name == "DataBlocks")
                    {
                        



                        PlcBlockComposition blockComposition = group.Blocks;



                        
                        MasterCopy masterCopySource = globalLibrary.MasterCopyFolder.Folders.Find("DataBlocks").MasterCopies.Find(dbName);




                       

                        
                            group.Blocks.CreateFrom(masterCopySource);
                            var db = group.Blocks.Find(masterCopySource.Name) as DataBlock;
                            changeDataBlock(db);
                            Console.WriteLine("Block " + group.Blocks.Last().Name + " Created in folder " + group.Name);

                        return 1;
                    }


                }
            }
            else
            {
                Console.WriteLine("There are no DataBlock Master Copies");
                return -1;
            }
            return -1;
        }




        //É possível mudar nome e número do datablock. Desenvolver mais tarde
        public void changeDataBlock(DataBlock db)
        {
            countDataBlocks();

            int numero = numeroDBsCylinder + 1;
            db.SetAttribute("Name", "Cilindro_" + numero);
            db.AutoNumber = false;
            db.SetAttribute("Number", numeroDBs+1);

            
            



        }


        //Com implementação de XML não será necessário
        public void importSingleFb()
        {
            MasterCopy masterCopy = globalLibrary.MasterCopyFolder.Folders.Find("FBs").MasterCopies.Find("Cylinders");
            var fb = plcSoftware.BlockGroup.Groups.Find("FBs").Blocks.CreateFrom(masterCopy);

            fb.Name = "Cylinders2";
            fb.AutoNumber = false;
            fb.Number = 3;

            Console.WriteLine(fb.IsConsistent);

            
        }


        public void copyToProjectLibraryPrompt()
        {

            projectLibrary = projectTIA.ProjectLibrary;
            bool existe = false;

            while (existe == false)
            {


                foreach (var folder in projectLibrary.TypeFolder.Folders)
                {
                    if (folder.Name == "HMI") existe = true;
                }

                if (existe == false)
                {
                    Console.WriteLine("Please Copy HMI Folder From Global Library to Project Library");

                    Console.WriteLine("Press Enter when Done...");
                    Console.ReadLine();
                }
            }
            
        }

       
        

        #endregion



        #region funções de screens e faceplates 

        //Cria um folder para Screens   (Retorna false se um folder com o mesmo nome não existia, criando um| Retorna true se folder com o mesmo nome já existia)
        public bool createScreenFolder(bool isUnified, string name = "Screens")
        {
            bool existe = false;

            int numScreenFolders;



            //Verifica se o folder já existe
            if (isUnified)
            {
                numScreenFolders = hmiSoftware.ScreenGroups.Count;
            }
            else
            {
                numScreenFolders = hmiTarget.ScreenFolder.Folders.Count;
            }
                

                if (numScreenFolders != 0)
                {



                for (int i = 0; i < numScreenFolders; i++)
                {
                    if (isUnified == false)
                    {
                        if (hmiTarget.ScreenFolder.Folders[i].Name == name)
                        {
                            existe = true;
                            Console.WriteLine("Folder " + name + " already exists");
                        }
                    }
                    else
                    {
                        if (hmiSoftware.ScreenGroups[i].Name == name)
                        {
                            existe = true;
                            Console.WriteLine("Folder " + name + " already exists");
                        }
                    }
                }
                    
                }

                if (existe == false && isUnified == false)
                {
                    hmiTarget.ScreenFolder.Folders.Create(name); //Cria o Folder
                    Console.WriteLine("Folder " + name + " created");
                }else if(existe == false && isUnified == true)
                {
                    hmiSoftware.ScreenGroups.Create(name);
                    Console.WriteLine("Folder " + name + " created");
                }





                return existe;

        }

        public void createScreen(bool isUnified , string folderName = "Screens", string screenName = "Screen_1")
        {
           bool existeFolder = false; //assume que não existe folder
           ScreenFolder screenFolder = null;
           Screen screen = null;
           HmiScreenGroup screenGroup;
           HmiScreen hmiScreen = null;
           int numScreenFolders;
           int screenFolderIndex = 0;

            if (hmiTarget != null && isUnified == false || hmiSoftware != null && isUnified == true)
                {


                if (isUnified)
                {
                    numScreenFolders = hmiSoftware.ScreenGroups.Count;
                    screenFolderIndex = 0;
                }
                else
                {
                    numScreenFolders = hmiTarget.ScreenFolder.Folders.Count;
                }
                    
                    
                    if (numScreenFolders == 0)
                    {
                        Console.WriteLine("Folder Doesn't exist");
                        createScreenFolder(isUnified,folderName);
                        
                    }else
                    {
                        for(int i = 0; i<numScreenFolders; i++)                         
                        {
                            if (isUnified)
                            {
                                
                                if (hmiSoftware.ScreenGroups[i].Name == folderName)
                                {
                                    screenGroup = hmiSoftware.ScreenGroups[i];
                                    existeFolder = true;
                                    screenFolderIndex = i;
                                
                                }
                            }
                            else
                            {
                                
                                if (hmiTarget.ScreenFolder.Folders[i].Name == folderName)   //
                                {

                                    screenFolder = hmiTarget.ScreenFolder.Folders[i];       //
                                    existeFolder = true;
                                
                                }                                                           //
                            }
                            
                        }

                        if(existeFolder == false)
                        {
                        Console.WriteLine("Folder Doesn't exist");
                        createScreenFolder(isUnified, folderName);
                        screenFolderIndex = numScreenFolders;
                        }



                        

                     }

                if (isUnified)
                {
                    numScreenFolders = hmiSoftware.ScreenGroups.Count();
                    
                    int numScreensInFolder = hmiSoftware.ScreenGroups[screenFolderIndex].Screens.Count();
                    bool existeScreen = false;
                    
                    for (int k = 0; k < numScreenFolders; k++)
                    {
                        numScreensInFolder = hmiSoftware.ScreenGroups[k].Screens.Count();

                        for (int j = 0; j < numScreensInFolder; j++)
                        {
                            if (hmiSoftware.ScreenGroups[k].Screens[j].Name == screenName)
                            {
                                existeScreen = true;
                                Console.WriteLine("Screen " + screenName + " already exists in " + folderName + " folder");
                            }
                        }
                    }
                    if(existeScreen == false)
                    {
                        hmiScreen = hmiSoftware.ScreenGroups[screenFolderIndex].Screens.Create(screenName);
                        Console.WriteLine("Screen " + hmiScreen.Name + " created in folder " + folderName);
                    }
                   
                }
                else
                {
                    MasterCopy masterCopy = globalLibrary.MasterCopyFolder.Folders.Find("Screens").MasterCopies[0];
                    screen = hmiTarget.ScreenFolder.Folders.Find(folderName).Screens.CreateFrom(masterCopy);


                    Console.WriteLine("Screen " + screen.Name + " created in folder " + folderName);
                }



            }
            else
            {
                Console.WriteLine("hmiTarget and hmiSoftware are null");
            }

            
        }


        public void getAllTemplatesFromLibrary(bool isUnified)
        {
            if (!isUnified)
            {
                for(int i = 0; i < globalLibrary.MasterCopyFolder.Folders.Find("Template").MasterCopies.Count; i++)
                {
                    getTemplateFomLibrary(globalLibrary.MasterCopyFolder.Folders.Find("Template").MasterCopies[i].Name);
                }
            }
        }
        public void getTemplateFomLibrary(string templateName)
        {
            bool existe = false;
            MasterCopy masterCopy = globalLibrary.MasterCopyFolder.Folders.Find("Template").MasterCopies.Find(templateName);

            for(int i = 0; i < hmiTarget.ScreenTemplateFolder.ScreenTemplates.Count; i++)
            {
                if (hmiTarget.ScreenTemplateFolder.ScreenTemplates[i].Name == templateName)
                {
                    existe = true;
                }
            }
            if(existe == false)
            {
                hmiTarget.ScreenTemplateFolder.ScreenTemplates.CreateFrom(masterCopy);
                Console.WriteLine("Template " + templateName + " added to template folder");
            }
            else
            {
                Console.WriteLine("Template " + templateName + " already exists in template folder");
            }
            



               



        }


        //Não Funciona
        public void createFaceplate(string folderName, string screenName, string faceplateName)
        {

            hmiSoftware.ScreenGroups.Find(folderName).Screens.Find(screenName).ScreenItems.Create<HmiFaceplateContainer>(faceplateName);

            var faceplate = hmiSoftware.ScreenGroups.Find(folderName).Screens.Find(screenName).ScreenItems[0] as HmiFaceplateContainer;
            var containedType = faceplate.ContainedType;

            

            
            hmiSoftware.ScreenGroups.Find(folderName).Screens.Find(screenName).ScreenItems[0].SetAttribute("Visible", true);
            
            




        }

        #endregion
        


        #region Importação de XML's
        //Importa XML de uma FB
        public void importFB()
        {
            var fbFolder = plcSoftware.BlockGroup.Groups.Find("FBs");



            FileInfo info = new FileInfo(string.Format(@"C:\Temp\teste2Openess\Cylinders.xml"));

            fbFolder.Blocks.Import(info, ImportOptions.Override);

            Console.WriteLine("FB Imported");




        }
        
        //Importa XML de uma FC
        public void importFC()
        {
            var fcFolder = plcSoftware.BlockGroup.Groups.Find("FCs");
            FileInfo info = new FileInfo(string.Format(@"C:\Temp\teste2Openess\Cylinders_write.xml"));
            fcFolder.Blocks.Import(info, ImportOptions.Override);

            //var bloco = fcFolder.Blocks.Find("Cylinders");

            //ICompilable serviceCompilable = bloco.GetService<ICompilable>();

            //serviceCompilable.Compile();


            Console.WriteLine("FC Imported");

        }

        //Importa XML de um DB
        public void importDB()
        {
            countDataBlocks();
            var fcFolder = plcSoftware.BlockGroup.Groups.Find("DataBlocks");

            FileInfo info = new FileInfo(string.Format(@"C:\Temp\teste2Openess\Cylinders_DB_write.xml"));

            fcFolder.Blocks.Import(info, ImportOptions.Override);

            

            
            
            Console.WriteLine("DB Imported");

            
        }

        public void changeMain()
        {
            var plcFolder = plcSoftware.BlockGroup.Groups;


            var mainBlock = plcSoftware.BlockGroup.Blocks;
            FileInfo info = new FileInfo(string.Format(@"C:\Temp\teste2Openess\Main.xml"));

            mainBlock.Import(info, ImportOptions.Override);

            Console.WriteLine("Main Block Imported");

        }


        public void importScreen()
        {
            FileInfo file = new FileInfo(string.Format(@"C:\Temp\teste2Openess\Screen_write.xml"));
            hmiTarget.ScreenFolder.Folders[0].Screens.Import(file, ImportOptions.Override);

            Console.WriteLine("Screen Imported");
        }

        #endregion



        #region Escrita do Documento XML de uma DB de Cilindros
        //Escreve Document Info no XML, usado em todos os objetos do TIA no PLC
        public void writeXmlDocumentInfo(XmlWriter writer)
        {
            writer.WriteStartElement("DocumentInfo");
            writer.WriteElementString("Created", "2025-03-19T15:18:57.2065646Z");
            writer.WriteElementString("ExportSetting", "WithDefaults, WithReadOnly");
            writer.WriteStartElement("InstalledProducts");
            writer.WriteStartElement("Product");
            writer.WriteElementString("DisplayName", "Totally Integrated Automation Portal");
            writer.WriteElementString("DisplayVersion", "V18");
            writer.WriteEndElement();
            writer.WriteStartElement("OptionPackage");
            writer.WriteElementString("DisplayName", "TIA Portal Openness");
            writer.WriteElementString("DisplayVersion", "V18");
            writer.WriteEndElement();
            writer.WriteStartElement("OptionPackage");
            writer.WriteElementString("DisplayName", "TIA Portal Version Control Interface");
            writer.WriteElementString("DisplayVersion", "V18");
            writer.WriteEndElement();
            writer.WriteStartElement("Product");
            writer.WriteElementString("DisplayName", "STEP 7 Professional");
            writer.WriteElementString("DisplayVersion", "V18");
            writer.WriteEndElement();
            writer.WriteStartElement("OptionPackage");
            writer.WriteElementString("DisplayName", "STEP 7 Safety");
            writer.WriteElementString("DisplayVersion", "V18");
            writer.WriteEndElement();
            writer.WriteStartElement("Product");
            writer.WriteElementString("DisplayName", "WinCC Advanced / Unified PC");
            writer.WriteElementString("DisplayVersion", "V18");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();
        }
        //Escreve uma estrutura semelhante a esta <CodeModifiedDate ReadOnly="true">2025-03-19T15:17:54.4622546Z</CodeModifiedDate>
        public void writeXmlElementWithAtribute(string elementString, string elementValue, string atributeString, string atributeValue, XmlWriter writer)
        {
            writer.WriteStartElement(elementString);
            writer.WriteAttributeString(atributeString, atributeValue);
            writer.WriteString(elementValue);
            writer.WriteEndElement();
        }
        public void writeXmlElementWithTwoAtributes(string elementString, string elementValue, string atributeString, string atributeValue, string secondAttributeString, string secondAttributeValue, XmlWriter writer)
        {
            writer.WriteStartElement(elementString);
            writer.WriteAttributeString(atributeString, atributeValue);
            writer.WriteAttributeString(secondAttributeString, secondAttributeValue);
            writer.WriteString(elementValue);
            writer.WriteEndElement();
        }

        public void writeXmlElementWithThreeAtributes(string elementString, string elementValue, string atributeString, string atributeValue, string secondAttributeString, string secondAttributeValue, string thirdAttribute, string thirdAttributeValue ,XmlWriter writer)
        {
            writer.WriteStartElement(elementString);
            writer.WriteAttributeString(atributeString, atributeValue);
            writer.WriteAttributeString(secondAttributeString, secondAttributeValue);
            writer.WriteAttributeString(thirdAttribute, thirdAttributeValue);
            writer.WriteString(elementValue);
            writer.WriteEndElement();
        }

        public void writeXmlMemberElement(string name, string dataType, XmlWriter writer)
        {
            writer.WriteStartElement("Member");
            writer.WriteAttributeString("Name", name);
            writer.WriteAttributeString("Datatype", dataType);
            writer.WriteEndElement();
        }
        public void writeXmlInterfaceDb(XmlWriter writer, int numCylindros)
        {   
            
            

                    for(int i = 0; i < numCylindros; i++)
            {
                string name = "Cilindro" + (i + 1);
                writer.WriteStartElement("Member");
                    writer.WriteAttributeString("Name", name);
                    writer.WriteAttributeString("Datatype", "\"CTRL_Cylinder\"");
                    writer.WriteAttributeString("Remanence", "NonRetain");
                    writer.WriteAttributeString("Accessibility", "Public");
                        writer.WriteStartElement("AttributeList");
                            writeXmlElementWithTwoAtributes("BooleanAttribute", "true", "Name", "ExternalAccessible", "SystemDefined", "true",  writer);
                            writeXmlElementWithTwoAtributes("BooleanAttribute", "true", "Name", "ExternalVisible", "SystemDefined", "true",  writer);
                            writeXmlElementWithTwoAtributes("BooleanAttribute", "true", "Name", "ExternalWritable", "SystemDefined", "true",  writer);
                            writeXmlElementWithThreeAtributes("BooleanAttribute", "true", "Name", "UserVisible","Informative", "true", "SystemDefined", "true",  writer);
                            writeXmlElementWithThreeAtributes("BooleanAttribute", "false", "Name", "UserReadOnly", "Informative", "true", "SystemDefined", "true",  writer);
                            writeXmlElementWithThreeAtributes("BooleanAttribute", "true", "Name", "UserDeletable", "Informative", "true", "SystemDefined", "true",  writer);
                            writeXmlElementWithTwoAtributes("BooleanAttribute","false", "Name", "SetPoint", "SystemDefined", "true",  writer);
                        writer.WriteEndElement();
                        writer.WriteStartElement("Sections");
                            writer.WriteStartElement("Section");
                            writer.WriteAttributeString("Name", "None");
                                writeXmlMemberElement("name", "String[20]", writer);
                                writer.WriteStartElement("Member");
                                writer.WriteAttributeString("Name", "Status");
                                writer.WriteAttributeString("Datatype", "\"CTRL_DeviceStatus\"");
                                    writer.WriteStartElement("Sections");
                                            writer.WriteStartElement("Section");
                                            writer.WriteAttributeString("Name", "None");
                                                writeXmlMemberElement("ready", "Bool", writer);
                                                writeXmlMemberElement("done", "Bool", writer);
                                                writeXmlMemberElement("busy", "Bool", writer);
                                                writeXmlMemberElement("idle", "Bool", writer);
                                                writeXmlMemberElement("nextDeviceReady", "Bool", writer);
                                                writeXmlMemberElement("error", "Bool", writer);
                                                writeXmlMemberElement("reset", "Bool", writer);
                                                writeXmlMemberElement("step", "Int", writer);
                                                writeXmlMemberElement("homeStep", "Int", writer);
                                                writeXmlMemberElement("manualMode", "Bool", writer);
                                                writeXmlMemberElement("homingOrder", "Bool", writer);
                                                writeXmlMemberElement("homed", "Bool", writer);
                                                writeXmlMemberElement("clock", "Bool", writer);
                                                writeXmlMemberElement("maximized", "Bool", writer);
                                             writer.WriteEndElement();
                                        writer.WriteEndElement();
                                writer.WriteEndElement();

                                writer.WriteStartElement("Member");
                                writer.WriteAttributeString("Name", "Enable");
                                writer.WriteAttributeString("Datatype", "Struct");
                                    writeXmlMemberElement("home", "Bool", writer);
                                    writeXmlMemberElement("work", "Bool", writer);
                                writer.WriteEndElement();

                                writer.WriteStartElement("Member");
                                writer.WriteAttributeString("Name", "Order");
                                writer.WriteAttributeString("Datatype", "Struct");
                                    writeXmlMemberElement("home", "Bool", writer);
                                    writeXmlMemberElement("work", "Bool", writer);
                                    writeXmlMemberElement("hmiHome", "Bool", writer);
                                    writeXmlMemberElement("hmiWork", "Bool", writer);
                                writer.WriteEndElement();
                                
                                writer.WriteStartElement("Member");
                                writer.WriteAttributeString("Name", "Time");
                                writer.WriteAttributeString("Datatype", "Struct");
                                    writeXmlMemberElement("filterHome", "Time", writer);
                                    writeXmlMemberElement("filterWork", "Time", writer);
                                    writeXmlMemberElement("timeoutHome", "Time", writer);
                                    writeXmlMemberElement("timeoutWork", "Time", writer);
                                writer.WriteEndElement();

                                writer.WriteStartElement("Member");
                                writer.WriteAttributeString("Name", "Sensor");
                                writer.WriteAttributeString("Datatype", "Struct");
                                    writeXmlMemberElement("home", "Bool", writer);
                                    writeXmlMemberElement("work", "Bool", writer);
                                writer.WriteEndElement();

                                writer.WriteStartElement("Member");
                                writer.WriteAttributeString("Name", "Output");
                                writer.WriteAttributeString("Datatype", "Struct");
                                    writeXmlMemberElement("home", "Bool", writer);
                                    writeXmlMemberElement("work", "Bool", writer);
                                writer.WriteEndElement();

                                writer.WriteStartElement("Member");
                                writer.WriteAttributeString("Name", "Position");
                                writer.WriteAttributeString("Datatype", "Struct");
                                    writeXmlMemberElement("home", "Bool", writer);
                                    writeXmlMemberElement("work", "Bool", writer);
                                writer.WriteEndElement();

                                writer.WriteStartElement("Member");
                                writer.WriteAttributeString("Name", "Error");
                                writer.WriteAttributeString("Datatype", "Struct");
                                    writeXmlMemberElement("home", "Bool", writer);
                                    writeXmlMemberElement("work", "Bool", writer);
                                writer.WriteEndElement();
                                
                                writer.WriteStartElement("Member");
                                writer.WriteAttributeString("Name", "hmiMaximized");
                                writer.WriteAttributeString("Datatype", "Struct");
                                    writeXmlMemberElement("errorHome", "Bool", writer);
                                    writeXmlMemberElement("errorWork", "Bool", writer);
                                    writeXmlMemberElement("error", "Bool", writer);
                                writer.WriteEndElement();
                                
                                writeXmlMemberElement("doesNotRetainOutput", "Bool", writer);

                            
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();

            }
                  

        }

        public void writeXmlfileDB(int numCylinders, string dbName)
        {
            countDataBlocks();
            FileInfo info = new FileInfo(string.Format(@"C:\Temp\teste2Openess\Cylinders_DB_teste.xml"));


            XmlWriter writer = XmlWriter.Create(@"C:\Temp\teste2Openess\Cylinders_DB_write.xml");

            writer.WriteStartDocument();
                writer.WriteStartElement("Document");
                    writer.WriteStartElement("Engineering");
                    writer.WriteAttributeString("version", "V18");
                    writer.WriteEndElement();
                    writeXmlDocumentInfo(writer);

            //Start Block Write
                    writer.WriteStartElement("SW.Blocks.GlobalDB");
                    writer.WriteAttributeString("ID", "0");
                        //Start Attribute List
                        writer.WriteStartElement("AttributeList");
                            writer.WriteElementString("AutoNumber", "true");
                            writeXmlElementWithAtribute("CodeModifiedDate", "2025-03-19T15:17:54.4622546Z", "ReadOnly", "true", writer);
                            writeXmlElementWithAtribute("CompileDate", "2025-03-19T15:18:53.1916012Z", "ReadOnly", "true", writer); 
                            writeXmlElementWithAtribute("CreationDate", "2025-03-19T15:17:54.4622546Z", "ReadOnly", "true", writer); 
                            writer.WriteElementString("DBAccessibleFromOPCUA", "true");
                            writer.WriteElementString("DBAccessibleFromWebserver", "true");
                            writeXmlElementWithAtribute("DownloadWithoutReinit", "false", "ReadOnly", "true", writer);
                            writer.WriteElementString("HeaderAuthor", "");
                            writer.WriteElementString("HeaderFamily", "");
                            writer.WriteElementString("HeaderName", "");
                            writer.WriteElementString("HeaderVersion", "0.1");

                            writer.WriteStartElement("Interface");
                                writer.WriteStartElement("Sections", "http://www.siemens.com/automation/Openness/SW/Interface/v5");
                                writer.WriteAttributeString("xmlns", "http://www.siemens.com/automation/Openness/SW/Interface/v5");
                                    writer.WriteStartElement("Section");
                                    writer.WriteAttributeString("Name", "Static");
                            writeXmlInterfaceDb(writer, numCylinders);

                                              writer.WriteEndElement();
                                
                            
                             
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writeXmlElementWithAtribute("InterfaceModifiedDate", "2025-03-19T15:17:54.4622546Z", "ReadOnly", "true", writer);
                            writeXmlElementWithAtribute("IsConsistent", "true", "ReadOnly", "true", writer);
                            writeXmlElementWithAtribute("IsKnowHowProtected", "false", "ReadOnly", "true", writer);
                            writer.WriteElementString("IsOnlyStoredInLoadMemory", "false");
                            writeXmlElementWithAtribute("IsPLCDB", "false", "ReadOnly", "true", writer);
                            writer.WriteElementString("IsRetainMemResEnabled", "false");
                            writer.WriteElementString("IsWriteProtectedInAS", "false");
                            writer.WriteElementString("MemoryLayout", "Optimized");
                            writer.WriteElementString("MemoryReserve", "100");
                            writeXmlElementWithAtribute("ModifiedDate", "2025-03-19T15:17:54.4622546Z", "ReadOnly", "true", writer);
                            writer.WriteElementString("Name", dbName);
                            writer.WriteElementString("Namespace", "");
                            string numDbString = (numeroDBs + 1).ToString();
                            writer.WriteElementString("Number", numDbString);
                            
                            writeXmlElementWithAtribute("ParameterModified", "2025-03-19T15:17:54.4622546Z", "ReadOnly", "true", writer);
                            writer.WriteElementString("ProgrammingLanguage", "DB");
                            writeXmlElementWithAtribute("StructureModified", "2025-03-19T15:17:54.4622546Z", "ReadOnly", "true", writer);
                            writer.WriteEndElement();
            
            //End Attribute List

            //Start Object List

            writer.WriteStartElement("ObjectList");
                        writer.WriteStartElement("MultilingualText");
                        writer.WriteAttributeString("ID", "1");
                        writer.WriteAttributeString("CompositionName", "Comment");
                                writer.WriteStartElement("ObjectList");
                                    writer.WriteStartElement("MultilingualTextItem");
                                    writer.WriteAttributeString("ID", "2");
                                    writer.WriteAttributeString("CompositionName", "Items");
                                                    writer.WriteStartElement("AttributeList");      
                                                        writer.WriteElementString("Culture", "en-US");
                                                        writer.WriteElementString("Text", "");
                                                    writer.WriteEndElement();
                                            writer.WriteEndElement();
                                    writer.WriteEndElement();
                               writer.WriteEndElement();
                         
                        
                        writer.WriteStartElement("MultilingualText");
                        writer.WriteAttributeString("ID", "3");
                        writer.WriteAttributeString("CompositionName", "Title");
                                writer.WriteStartElement("ObjectList");
                                    writer.WriteStartElement("MultilingualTextItem");
                                    writer.WriteAttributeString("ID", "4");
                                    writer.WriteAttributeString("CompositionName", "Items");
                                                    writer.WriteStartElement("AttributeList");      
                                                        writer.WriteElementString("Culture", "en-US");
                                                        writer.WriteElementString("Text", "");
                                                    writer.WriteEndElement();
                                            writer.WriteEndElement();
                                    writer.WriteEndElement();
                               writer.WriteEndElement();
            writer.WriteEndElement();
            



            //End Block Write
            writer.WriteEndDocument();
            writer.Flush();
            writer.Close();
            
            

            

            
        }

        #endregion Escrita do Documento XML de uma DB de Cilindros


        #region Escrita do Documento XML de uma FC de Cilindros

        /// <summary>
        /// Terminar
        /// </summary>
        /// <param name="writer"></param>
        /// 

        public void writeXmlSingleWire(XmlWriter writer, int Uid_1, int Uid_2, int Uid_3, string name)
        {
            writer.WriteStartElement("Wire");
            writer.WriteAttributeString("UId", Uid_1.ToString());
                writer.WriteStartElement("OpenCon");
                writer.WriteAttributeString("UId", Uid_2.ToString());
                writer.WriteEndElement();
                writer.WriteStartElement("NameCon");
                writer.WriteAttributeString("UId", Uid_3.ToString());
                writer.WriteAttributeString("Name", name);
                writer.WriteEndElement();
            writer.WriteEndElement();
        }
        public void writeXmlWires(XmlWriter writer)
        {
            writer.WriteStartElement("Wire");
                writer.WriteAttributeString("UId", "42");
                writer.WriteStartElement("Powerrail");
                writer.WriteEndElement();
                writer.WriteStartElement("NameCon");
                writer.WriteAttributeString("UId", "22");
                writer.WriteAttributeString("Name", "en");
                writer.WriteEndElement();
            writer.WriteEndElement();


            writeXmlSingleWire(writer, 43, 24, 22, "name");
            writeXmlSingleWire(writer, 44, 25, 22, "enableHome");
            writeXmlSingleWire(writer, 45, 26, 22, "enableWork");
            writeXmlSingleWire(writer, 46, 27, 22, "doorOpen");
            writeXmlSingleWire(writer, 47, 28, 22, "manualMode");
            writeXmlSingleWire(writer, 48, 29, 22, "reset");
            writeXmlSingleWire(writer, 49, 30, 22, "iHome");
            writeXmlSingleWire(writer, 50, 31, 22, "iWork");
            writeXmlSingleWire(writer, 51, 32, 22, "orderHome");
            writeXmlSingleWire(writer, 52, 33, 22, "orderWork");
            writeXmlSingleWire(writer, 53, 34, 22, "doesNotRetainOutput");
            writeXmlSingleWire(writer, 54, 35, 22, "timeFilterHome");
            writeXmlSingleWire(writer, 55, 36, 22, "timeFilterWork");
            writeXmlSingleWire(writer, 56, 37, 22, "timeTimeout");

            writer.WriteStartElement("Wire");
            writer.WriteAttributeString("UId", "57");
                writer.WriteStartElement("IdentCon");
                writer.WriteAttributeString("UId", "21");
                writer.WriteEndElement();
                writer.WriteStartElement("NameCon");
                writer.WriteAttributeString("UId", "22");
                writer.WriteAttributeString("Name", "Cylinder");
                writer.WriteEndElement();
            writer.WriteEndElement();


            
            writer.WriteStartElement("Wire");
            writer.WriteAttributeString("UId", "58");
                writer.WriteStartElement("NameCon");
                writer.WriteAttributeString("UId", "22");
                writer.WriteAttributeString("Name", "outputHome");
                writer.WriteEndElement();
                writer.WriteStartElement("OpenCon");
                writer.WriteAttributeString("UId", "38");
                writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("Wire");
            writer.WriteAttributeString("UId", "59");
                writer.WriteStartElement("NameCon");
                writer.WriteAttributeString("UId", "22");
                writer.WriteAttributeString("Name", "outputWork");
                writer.WriteEndElement();
                writer.WriteStartElement("OpenCon");
                writer.WriteAttributeString("UId", "39");
                writer.WriteEndElement();
            writer.WriteEndElement();


            writer.WriteStartElement("Wire");
            writer.WriteAttributeString("UId", "60");
                writer.WriteStartElement("NameCon");
                writer.WriteAttributeString("UId", "22");
                writer.WriteAttributeString("Name", "errorTimeoutWork");
                writer.WriteEndElement();
                writer.WriteStartElement("OpenCon");
                writer.WriteAttributeString("UId", "40");
                writer.WriteEndElement();
            writer.WriteEndElement();


            writer.WriteStartElement("Wire");
            writer.WriteAttributeString("UId", "61");
                writer.WriteStartElement("NameCon");
                writer.WriteAttributeString("UId", "22");
                writer.WriteAttributeString("Name", "errorTimeoutHome");
                writer.WriteEndElement();
                writer.WriteStartElement("OpenCon");
                writer.WriteAttributeString("UId", "41");
                writer.WriteEndElement();
            writer.WriteEndElement();




        }
        public void writeXmlPartParameter(string name, string section, string type, XmlWriter writer)
        {
            writer.WriteStartElement("Parameter");
            writer.WriteAttributeString("Name", name);
            writer.WriteAttributeString("Section", section);
            writer.WriteAttributeString("Type", type);
                writeXmlElementWithTwoAtributes("StringAttribute", "S7_Visible", "Name", "InterfaceFlags", "Informative", "true", writer);
            writer.WriteEndElement();
        }

        
        public string intToHex(int idCounter)
        {
            string hexVal = Convert.ToString(idCounter, 16);

            hexVal.ToUpper();

            return hexVal;
        }
        public int writeXmlNetorksFc(int idCounter, XmlWriter writer, int numCylinders)
        {
            
            string numCilindroString;
            string bitOffset;
            for(int i = 0; i < numCylinders; i++)
            {
                bitOffset = (32 + 576 * i).ToString();
                writer.WriteStartElement("SW.Blocks.CompileUnit");
                
                writer.WriteAttributeString("ID", intToHex(idCounter));
                idCounter++;
                writer.WriteAttributeString("CompositionName", "CompileUnits");
                    writer.WriteStartElement("AttributeList");
                        writer.WriteStartElement("NetworkSource");
                            writer.WriteStartElement("FlgNet", "http://www.siemens.com/automation/Openness/SW/NetworkSource/FlgNet/v4");
                            writer.WriteAttributeString("xmlns", "http://www.siemens.com/automation/Openness/SW/NetworkSource/FlgNet/v4");
                                //Start Parts
                                writer.WriteStartElement("Parts");
                                    writer.WriteStartElement("Access");
                                    writer.WriteAttributeString("Scope", "GlobalVariable");
                                    writer.WriteAttributeString("UId", "21");
                                        writer.WriteStartElement("Symbol");
                                            writer.WriteStartElement("Component");
                                            writer.WriteAttributeString("Name", "Cylinders_DB");
                                            writer.WriteEndElement();
                                            numCilindroString = "Cilindro" + (i+1);
                                            writer.WriteStartElement("Component");
                                            writer.WriteAttributeString("Name", numCilindroString);
                                            writer.WriteEndElement();
                                            writer.WriteStartElement("Address");
                                            writer.WriteAttributeString("Area", "None");
                                            writer.WriteAttributeString("Type", "CTRL_Cylinder");
                                            string blockNumber = plcSoftware.BlockGroup.Groups.Find("DataBlocks").Blocks.Find("Cylinders_DB").Number.ToString();
                                            writer.WriteAttributeString("BlockNumber", blockNumber);
                                            writer.WriteAttributeString("BitOffset", bitOffset);
                                            writer.WriteAttributeString("Informative", "true");
                                            
                                            writer.WriteEndElement();
                                        writer.WriteEndElement();
                                    writer.WriteEndElement();
                                    

                                    writer.WriteStartElement("Call");
                                    writer.WriteAttributeString("UId", "22");
                                        writer.WriteStartElement("CallInfo");
                                        writer.WriteAttributeString("Name", "FB_Cylinder");
                                        writer.WriteAttributeString("BlockType", "FB");
                                            string fbBlockNumber = plcSoftware.BlockGroup.Groups.Find("FBs").Blocks.Find("FB_Cylinder").Number.ToString();
                                            writeXmlElementWithTwoAtributes("IntegerAttribute", fbBlockNumber, "Name", "BlockNumber",  "Informative", "true", writer);
                                            writeXmlElementWithTwoAtributes("DateAttribute", "2024-07-16T16:22:51", "Name", "ParameterModifiedTS",  "Informative", "true", writer);
                                                writer.WriteStartElement("Instance");
                                                writer.WriteAttributeString("Scope", "GlobalVariable");
                                                writer.WriteAttributeString("UId", "23");
                                                    numCilindroString = ("Cilindro_" + (i+1)).ToString();
                                                    writer.WriteStartElement("Component");
                                                    writer.WriteAttributeString("Name", numCilindroString);
                                                    writer.WriteEndElement();
                                                    
                                                    writer.WriteStartElement("Address");
                                                    blockNumber = plcSoftware.BlockGroup.Groups.Find("DataBlocks").Blocks.Find(numCilindroString).Number.ToString("");
                                                    writer.WriteAttributeString("Area", "DB");
                                                    writer.WriteAttributeString("Type", "FB_Cylinder");
                                                    writer.WriteAttributeString("BlockNumber", blockNumber);
                                                    writer.WriteAttributeString("BitOffset", "0");
                                                    writer.WriteAttributeString("Informative", "true");
                                                writer.WriteEndElement();
                                                writer.WriteEndElement();
                                                writeXmlPartParameter("name", "Input", "String[20]", writer);
                                                writeXmlPartParameter("enableHome", "Input", "Bool", writer);
                                                writeXmlPartParameter("enableWork", "Input", "Bool", writer);
                                                writeXmlPartParameter("doorOpen", "Input", "Bool", writer);
                                                writeXmlPartParameter("manualMode", "Input", "Bool", writer);
                                                writeXmlPartParameter("reset", "Input", "Bool", writer);
                                                writeXmlPartParameter("iHome", "Input", "Bool", writer);
                                                writeXmlPartParameter("iWork", "Input", "Bool", writer);
                                                writeXmlPartParameter("orderHome", "Input", "Bool", writer);
                                                writeXmlPartParameter("orderWork", "Input", "Bool", writer);
                                                writeXmlPartParameter("doesNotRetainOutput", "Input", "Bool", writer);
                                                writeXmlPartParameter("timeFilterHome", "Input", "Time", writer);
                                                writeXmlPartParameter("timeFilterWork", "Input", "Time", writer);
                                                writeXmlPartParameter("timeTimeout", "Input", "Time", writer);
                                                writeXmlPartParameter("outputHome", "Output", "Bool", writer);
                                                writeXmlPartParameter("outputWork", "Output", "Bool", writer);
                                                writeXmlPartParameter("errorTimeoutWork", "Output", "Bool", writer);
                                                writeXmlPartParameter("errorTimeoutHome", "Output", "Bool", writer);
                                                writeXmlPartParameter("Cylinder", "InOut", "\"CTRL_Cylinder\"", writer);

                                            writer.WriteEndElement();
                                    writer.WriteEndElement();
                                    writer.WriteEndElement();
                                    
                                    
                                
                                //End Parts


                                //Start Wires
                                writer.WriteStartElement("Wires");
                                writeXmlWires(writer);
                                    
                                //writeXmlWire();
                                    
                                writer.WriteEndElement();
                                writer.WriteEndElement();
                                //End Wires
                                writer.WriteEndElement();
                                writer.WriteElementString("ProgrammingLanguage", "LAD");
                                writer.WriteEndElement();
                                //Start Object List
                                writer.WriteStartElement("ObjectList");
                                
                                    writer.WriteStartElement("MultilingualText");
                                    writer.WriteAttributeString("ID", intToHex(idCounter));
                                    idCounter++;
                                    writer.WriteAttributeString("CompositionName", "Comment");
                                        writer.WriteStartElement("ObjectList");
                                            writer.WriteStartElement("MultilingualTextItem");
                                            writer.WriteAttributeString("ID", intToHex(idCounter));
                                            idCounter++;
                                            writer.WriteAttributeString("CompositionName", "Items");
                                                writer.WriteStartElement("AttributeList");
                                                    writer.WriteElementString("Culture", "en-US");
                                                    writer.WriteElementString("Text", "");
                                            writer.WriteEndElement();
                                        writer.WriteEndElement();
                                    writer.WriteEndElement();
                                    writer.WriteEndElement();
                                    writer.WriteStartElement("MultilingualText");
                                    writer.WriteAttributeString("ID", intToHex(idCounter));
                                    idCounter++;
                                    writer.WriteAttributeString("CompositionName", "Title");
                                        writer.WriteStartElement("ObjectList");
                                            writer.WriteStartElement("MultilingualTextItem");
                                            writer.WriteAttributeString("ID", intToHex(idCounter));
                                            idCounter++;
                                            writer.WriteAttributeString("CompositionName", "Items");
                                                writer.WriteStartElement("AttributeList");
                                                    writer.WriteElementString("Culture", "en-US");
                                                    string cilindroStr = "Cilindro " + (i+1);
                                                    writer.WriteElementString("Text", cilindroStr);
                                            writer.WriteEndElement();
                                        writer.WriteEndElement();
                                    writer.WriteEndElement();
                                
                                //End Object List


                            
                    
                        
                        
                        
                        writer.WriteEndElement();

                        writer.WriteEndElement();

                writer.WriteEndElement();

            }

            return idCounter;
        }

        public void writeXmlInterfaceFc(XmlWriter writer)
        {
            writer.WriteStartElement("Interface");
                writer.WriteStartElement("Sections", "http://www.siemens.com/automation/Openness/SW/Interface/v5");
                writer.WriteAttributeString("xmlns", "http://www.siemens.com/automation/Openness/SW/Interface/v5");
                    writer.WriteStartElement("Section");
                    writer.WriteAttributeString("Name", "Input");
                    writer.WriteEndElement();
                    writer.WriteStartElement("Section");
                    writer.WriteAttributeString("Name", "Output");
                    writer.WriteEndElement();
                    writer.WriteStartElement("Section");
                    writer.WriteAttributeString("Name", "InOut");
                    writer.WriteEndElement();
                    writer.WriteStartElement("Section");
                    writer.WriteAttributeString("Name", "Temp");
                    writer.WriteEndElement();
                    writer.WriteStartElement("Section");
                    writer.WriteAttributeString("Name", "Constant");
                    writer.WriteEndElement();

                    writer.WriteStartElement("Section");
                    writer.WriteAttributeString("Name", "Return");
                    
                        writer.WriteStartElement("Member");
                        writer.WriteAttributeString("Name", "Ret_Val");
                        writer.WriteAttributeString("Datatype", "Void");
                        writer.WriteAttributeString("Accessibility", "Public");
                            writer.WriteStartElement("AttributeList");
                            writer.WriteEndElement();
                        writer.WriteEndElement();
                    writer.WriteEndElement();

            writer.WriteEndElement();
            writer.WriteEndElement();
        }
        public void writeXmlFileFC(int numCylinders, string fcName)
        {
            FileInfo info = new FileInfo(string.Format(@"C:\Temp\teste2Openess\Cylinders.xml"));

            countFCs();
            XmlWriter writer = XmlWriter.Create(@"C:\Temp\teste2Openess\Cylinders_write.xml");

            int idCounter = 0;

            writer.WriteStartDocument();
                writer.WriteStartElement("Document");
                    writer.WriteStartElement("Engineering");
                    writer.WriteAttributeString("version", "V18");
                    writer.WriteEndElement();
                    writeXmlDocumentInfo(writer);
                    
                    writer.WriteStartElement("SW.Blocks.FC");
                    writer.WriteAttributeString("ID", "0");
                    idCounter++;
                        //Start Attribute List
                        writer.WriteStartElement("AttributeList");
                            writer.WriteElementString("AutoNumber", "true");
                            writeXmlElementWithAtribute("CodeModifiedDate", "2025-03-19T15:17:54.4622546Z", "ReadOnly", "true", writer);
                            writeXmlElementWithAtribute("CompileDate", "2025-03-19T15:18:53.1916012Z", "ReadOnly", "true", writer); 
                            writeXmlElementWithAtribute("CreationDate", "2025-03-19T15:17:54.4622546Z", "ReadOnly", "true", writer); 
                            writeXmlElementWithAtribute("HandleErrorsWithinBlock", "false", "ReadOnly", "true", writer);
                            writer.WriteElementString("HeaderAuthor", "");
                            writer.WriteElementString("HeaderFamily", "");
                            writer.WriteElementString("HeaderName", "");
                            writer.WriteElementString("HeaderVersion", "0.1");

                            //Start Interface
                            writeXmlInterfaceFc(writer);
                            //End Interface

                            writeXmlElementWithAtribute("InterfaceModifiedDate", "2025-03-19T15:17:54.4622546Z", "ReadOnly", "true", writer);
                            writeXmlElementWithAtribute("IsConsistent", "true", "ReadOnly", "true", writer);
                            writer.WriteElementString("IsIECCheckEnabled", "false");
                            writeXmlElementWithAtribute("IsKnowHowProtected", "false", "ReadOnly", "true", writer);
                            writeXmlElementWithAtribute("IsWriteProtected", "false", "ReadOnly", "true", writer);
                            writeXmlElementWithAtribute("LibraryConformanceStatus", "Error: The block contains calls of single instances. Warning: The object contains access to global data blocks.", "ReadOnly", "true", writer);
                            writer.WriteElementString("MemoryLayout", "Optimized");
                            writeXmlElementWithAtribute("ModifiedDate", "2025-03-19T15:17:54.4622546Z", "ReadOnly", "true", writer);
                            writer.WriteElementString("Name", fcName);
                            writer.WriteElementString("Namespace", "");
                            string fcNumberString = (numeroFCs +1).ToString();
                            writer.WriteElementString("Number", fcNumberString);
                            
                            writeXmlElementWithAtribute("ParameterModified", "2025-03-19T15:17:54.4622546Z", "ReadOnly", "true", writer);
                            writeXmlElementWithAtribute("PLCSimAdvancedSupport", "false", "ReadOnly", "true", writer);
                            writer.WriteElementString("ProgrammingLanguage", "LAD");
                            writer.WriteElementString("SetENOAutomatically", "false");
                            writeXmlElementWithAtribute("StructureModified", "2025-03-21T14:22:32.6241053Z", "ReadOnly", "true", writer);
                            writer.WriteElementString("UDABlockProperties", "");
                            writer.WriteElementString("UDAEnableTagReadback", "false");
                         writer.WriteEndElement();
                         //End Attribute List

                         //Start Object List
                         writer.WriteStartElement("ObjectList");
                            writer.WriteStartElement("MultilingualText");
                            writer.WriteAttributeString("ID", "1");
                            writer.WriteAttributeString("CompositionName", "Comment");
                                writer.WriteStartElement("ObjectList");
                                    writer.WriteStartElement("MultilingualTextItem");
                                    writer.WriteAttributeString("ID", "2");
                                    writer.WriteAttributeString("CompositionName", "Items");
                                                    writer.WriteStartElement("AttributeList");      
                                                        writer.WriteElementString("Culture", "en-US");
                                                        writer.WriteElementString("Text", "");
                                                    writer.WriteEndElement();
                                            writer.WriteEndElement();
                                    writer.WriteEndElement();
                               writer.WriteEndElement();
                        idCounter = 3;

                         //Start Networks
                         idCounter = writeXmlNetorksFc(idCounter, writer, numCylinders);
                            

                        //Start Final MultilingualText 
                        writer.WriteStartElement("MultilingualText");
                        writer.WriteAttributeString("ID", intToHex(idCounter));
                        idCounter++;
                        writer.WriteAttributeString("CompositionName", "Title");
                            writer.WriteStartElement("ObjectList");
                                writer.WriteStartElement("MultilingualTextItem");
                                writer.WriteAttributeString("ID", intToHex(idCounter));
                                idCounter++;
                                writer.WriteAttributeString("CompositionName", "Items");
                                    writer.WriteStartElement("AttributeList");
                                        writer.WriteElementString("Culture", "en-US");
                                        writer.WriteElementString("Text", "");
                                    writer.WriteEndElement();

                                writer.WriteEndElement();
                            writer.WriteEndElement();
                        writer.WriteEndElement();
                         
                         writer.WriteEndElement();
                         //End Object List

            writer.WriteEndElement();
            writer.WriteEndDocument();
            writer.Flush();
            writer.Close();
        }





        #endregion Escrita do Documento XML de uma FC de Cilindros


        # region Escrita do Documento XML de uma Screen com Cilindros

        

        public int writeFaceplateInstances(XmlWriter writer, int idCounter, int numCylindros )
        {
            string top;
            string left;
            writer.WriteStartElement("ObjectList");
            for(int i = 0; i < numCylindros; i++)
            {
                if (i < 2) {
                    top = "200";
                }
                else
                {
                    top = "460";
                }

                if(i%2 != 0)
                {
                    left = "140";
                }
                else
                {
                    left = "490";
                }
                writer.WriteStartElement("Hmi.Screen.FaceplateInstance");
                writer.WriteAttributeString("ID", intToHex(idCounter));
                idCounter++;
                writer.WriteAttributeString("CompositionName", "ScreenItems");
                    writer.WriteStartElement("AttributeList");
                        writer.WriteElementString("FaceplateTypeName", "HMI@$@Cylinder V 0.1.36");
                        writer.WriteElementString("Height","137");
                        writer.WriteElementString("Left", left);
                        string nameFP = "Cylinder_" + (i+1);
                        writer.WriteElementString("ObjectName",nameFP );
                        writer.WriteElementString("Resizing", "FixedSize");
                        writer.WriteElementString("TabIndex", (i+1).ToString());
                        writer.WriteElementString("Top", top);
                        writer.WriteElementString("Width", "249");
                    writer.WriteEndElement();

                    writer.WriteStartElement("ObjectList");
                    

                        writer.WriteStartElement("Hmi.Screen.InterfacePropertySimple");
                        writer.WriteAttributeString("ID", intToHex(idCounter));
                        idCounter++;
                        writer.WriteAttributeString("CompositionName", "Interface");
                            writer.WriteStartElement("AttributeList");
                                writer.WriteElementString("Name", "Cylinder");
                            writer.WriteEndElement();
                            writer.WriteStartElement("ObjectList");
                                writer.WriteStartElement("Hmi.Dynamic.TagConnectionDynamic");
                                writer.WriteAttributeString("ID", intToHex(idCounter));
                                idCounter += 2;
                                
                                writer.WriteAttributeString("CompositionName", "Dynamic");
                                    writer.WriteStartElement("AttributeList");
                                        writer.WriteElementString("Indirect", "false");
                                    writer.WriteEndElement();
                                    writer.WriteStartElement("LinkList");
                                        writer.WriteStartElement("Tag");
                                        writer.WriteAttributeString("TargetID", "@OpenLink");
                                            string cilindroAlvo = "Cylinders_DB_Cilindro" + (i+1);
                                            writer.WriteElementString("Name", cilindroAlvo);
                                        writer.WriteEndElement();
                                    writer.WriteEndElement();
                            
                                writer.WriteEndElement();
                            writer.WriteEndElement();

                        writer.WriteEndElement();
                    writer.WriteEndElement();

                writer.WriteEndElement();
            }

            writer.WriteEndElement();

            return idCounter;
        }
        public void writeXmlFileScreen(int numCilindros, string name)
        {

            

            XmlWriter writer = XmlWriter.Create(@"C:\Temp\teste2Openess\Screen_write.xml");

            int idCounter = 0;

            writer.WriteStartDocument();
                writer.WriteStartElement("Document");
                    writer.WriteStartElement("Engineering");
                    writer.WriteAttributeString("version", "V18");
                    writer.WriteEndElement();
                    writeXmlDocumentInfo(writer);
                    

                    //Start Screen Parameters

                    writer.WriteStartElement("Hmi.Screen.Screen");
                    writer.WriteAttributeString("ID", "0");
                    idCounter++;
                        writer.WriteStartElement("AttributeList");
                            writer.WriteElementString("ActiveLayer", "0");
                            writer.WriteElementString("BackColor", "182, 182, 182");
                            writer.WriteElementString("GridColor", "0, 0, 0");
                            writer.WriteElementString("Height", "800");
                            writer.WriteElementString("Name", name);
                            writer.WriteElementString("Number", "1");
                            writer.WriteElementString("Visible", "true");
                            writer.WriteElementString("Width", "1280");
                        writer.WriteEndElement();

                      //Start Object List
                      writer.WriteStartElement("ObjectList");

                        writer.WriteStartElement("MultilingualText");
                        writer.WriteAttributeString("ID", "1");
                        idCounter++;
                        writer.WriteAttributeString("CompositionName", "HelpText");

                            writer.WriteStartElement("ObjectList");
                                writer.WriteStartElement("MultilingualTextItem");
                                writer.WriteAttributeString("ID", "2");
                                idCounter++;
                                writer.WriteAttributeString("CompositionName", "Items");
                                    writer.WriteStartElement("AttributeList");
                                        writer.WriteElementString("Culture", "en-US");
                                        writer.WriteElementString("Text", "");
                                    writer.WriteEndElement();
                                writer.WriteEndElement();
                            writer.WriteEndElement();
                        writer.WriteEndElement();

                        
                        writer.WriteStartElement("Hmi.Screen.ScreenLayer");
                        writer.WriteAttributeString("ID", "3");
                        idCounter++;
                        writer.WriteAttributeString("CompositionName", "Layers");
                            writer.WriteStartElement("AttributeList");
                                writer.WriteElementString("Index", "0");
                                writer.WriteElementString("Name", "");
                                writer.WriteElementString("VisibleES", "true");
                            writer.WriteEndElement();

                            
                            
                        writeFaceplateInstances(writer, idCounter, 4);
                        writer.WriteEndElement();
                      
                      
                      writer.WriteEndElement();
                      //End Object List

                

                    writer.WriteEndElement();

                    //End Screen Parameters

                writer.WriteEndElement();


            writer.WriteEndDocument();
            writer.Flush();
            writer.Close();

        }

        #endregion



        #region tag table da hmi

        
        
        public void exportTagTable()
        {
            FileInfo file = new FileInfo(string.Format(@"C:\Temp\teste2Openess\tagTable.xml"));

            hmiTarget.TagFolder.TagTables.Find("CylinderTT").Export(file, ExportOptions.WithDefaults);


        }


        public void importTagTable()
        {
            FileInfo file = new FileInfo(string.Format(@"C:\Temp\teste2Openess\tagTable.xml"));
            hmiTarget.TagFolder.TagTables.Import(file, ImportOptions.Override);
        }


        #endregion tag table da hmi


        #region Escrita de um documento em XML de uma TagTable de HMI

        public void writeXmlTagTableCilindro(int numCilindros, string name)
        {
            XmlWriter writer = XmlWriter.Create(@"C:\Temp\teste2Openess\TagTable_write.xml");

            int idCounter = 0;


            writer.WriteStartDocument();
                writer.WriteStartElement("Document");
                    writer.WriteStartElement("Engineering");
                    writer.WriteAttributeString("version", "V18");
                    writer.WriteEndElement();
                    writeXmlDocumentInfo(writer);

                    //Start Tag Table

                    writer.WriteStartElement("Hmi.Tag.TagTable");
                    writer.WriteAttributeString("ID", "0");
                    idCounter++;
                        writer.WriteStartElement("AttributeList");
                            writer.WriteElementString("Name", name);
                        writer.WriteEndElement();

                    //Start Object List (Tag Table Members)
                        writer.WriteStartElement("ObjectList");


                        writer.WriteEndElement();
                    //End Object List (Tag Table Members)
                    writer.WriteEndElement();
                writer.WriteEndElement();
            writer.WriteEndDocument();
            writer.Flush();
            writer.Close();
        }

        #endregion










    }



}


