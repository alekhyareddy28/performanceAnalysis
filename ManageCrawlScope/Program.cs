using System;
using Microsoft.Search.Interop;
[assembly: CLSCompliant(true)]
//[assembly: OleDbPermission(SecurityAction.RequestMinimum, Unrestricted = true)]

namespace WindowsSearch
{

    class Program
    { 
        [STAThread]
        static void Main(string[] args)
        {
            // This uses the Microsoft.Search.Interop assembly
            CSearchManager manager = new CSearchManager();

            // SystemIndex catalog is the default catalog in Windows
            ISearchCatalogManager catalogManager = manager.GetCatalog("SystemIndex");

            // Get the ISearchQueryHelper which will help us to translate AQS --> SQL necessary to query the indexer
            ISearchCrawlScopeManager crawlScopeManager = catalogManager.GetCrawlScopeManager();
            
            // Function which displays all the search roots
            DisplaySearchRoots(crawlScopeManager);

            // Display Scope rules
            DisplayScopeRules(crawlScopeManager);

            // Function to remove a search root
            string pathToRemove = @"file:///E:\";
            crawlScopeManager.RemoveRoot(pathToRemove);

            // to verify that it was deleted
            DisplaySearchRoots(crawlScopeManager);
            DisplayScopeRules(crawlScopeManager);

            CSearchRoot addSearchRoot = new CSearchRoot();
            addSearchRoot.RootURL = pathToRemove;
            // to add search root
            crawlScopeManager.AddRoot(addSearchRoot);
            // to add scope rules
            crawlScopeManager.AddUserScopeRule(pathToRemove, 1, 1, 1);

            // to verify that it was added again
            DisplaySearchRoots(crawlScopeManager);
            DisplayScopeRules(crawlScopeManager);

            // To commit all the changes to the windows search indexer, without this line of code it does not get modified in the global windows indexer
            crawlScopeManager.SaveAll();

            Console.Read();
        }
        

        // Function enumerates all the scope rules
        private static void DisplayScopeRules(ISearchCrawlScopeManager crawlScopeManager)
        {
            //To search for the scope rules
            IEnumSearchScopeRules scopeRules = crawlScopeManager.EnumerateScopeRules();
            CSearchScopeRule scopeRule;
            uint numScopes = 0;

            bool nextExists = true;
            while (nextExists)
            {
                try
                {
                    scopeRules.Next(1, out scopeRule, ref numScopes);
                    Console.WriteLine("Scope Rule " + scopeRule.PatternOrURL);
                    /*Console.WriteLine(numScopes);*/
                }
                catch (Exception e)
                {
                    nextExists = false;
                }
            }
        }

        // function which displays all the search roots
        private static void DisplaySearchRoots(ISearchCrawlScopeManager crawlScopeManager)
        {
            // Enumerates all the roots
            IEnumSearchRoots searchRoots = crawlScopeManager.EnumerateRoots();
            CSearchRoot searchRoot;
            uint numElements = 0;

            bool nextExists = true;
            while (nextExists)
            {
                try
                {
                    searchRoots.Next(1, out searchRoot, ref numElements);
                    Console.WriteLine("Search Root " + searchRoot.RootURL);
                    /*Console.WriteLine(numElements);*/
                }
                catch (Exception e)
                {
                    nextExists = false;
                }
            }
        }
    }
}