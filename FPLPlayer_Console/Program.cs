using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using FPLPlayer_Console.Models;
using FPLPlayer_Console.Utils;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace FPLPlayer_Console
{
    class Program
    {
        static void Main(string[] args)
        {
            //List<RootObject> playerList = new List<RootObject>();
            //List<HistorySummary> playerStatList = new List<HistorySummary>();
            List<History> playerStatList = new List<History>();
            List<Element> playerDataList = new List<Element>();
            List<Team> teamDataList = new List<Team>();
            GetPlayerStats(playerStatList);
            playerDataList = GetPlayerData();
            teamDataList = GetTeamData();
            ExcelOperations.PopulateResults(playerDataList, playerStatList, teamDataList);

            Console.WriteLine("Operation Completed");
            Console.ReadKey();
        }

        private static void GetPlayerStats(List<History> playerStatList)
        {
            WebClient client = new WebClient();
            client.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36");
            for (int i = 1; i < 609; i++)
            {
                try
                {
                    string url = "https://fantasy.premierleague.com/drf/element-summary/";
                    string response = client.DownloadString(url + i);
                    if (string.IsNullOrEmpty(response))
                    {
                        Console.WriteLine("No data for i = " + i);
                    }
                    else
                    {
                        //Console.WriteLine("Completd for player : " + i);
                    }
                    RootObject ro = JsonConvert.DeserializeObject<RootObject>(response);
                    Console.WriteLine("Player " + i);
                    //foreach (HistorySummary hSummary in ro.history_summary)
                    //{
                    //    playerStatList.Add(hSummary);
                    //}
                    foreach (History history in ro.history)
                    {
                        playerStatList.Add(history);
                    }
                }
                catch (Exception ex)
                {
                    string exceptionMessage = ex.Message;
                    Console.WriteLine(exceptionMessage);
                }
            }
        }

        private static List<Element> GetPlayerData()
        {
            string playerDataJSON = System.IO.File.ReadAllText(@"C:\Users\sbekal\Dropbox\FPL\Player_JSON_week30.json");
            PlayerData playerDataResponse = JsonConvert.DeserializeObject<PlayerData>(playerDataJSON);
            return playerDataResponse.elements;
        }

        private static List<Team> GetTeamData()
        {
            string teamDataJSON = System.IO.File.ReadAllText(@"C:\Users\sbekal\Dropbox\FPL\Team_Info.json");
            TeamInfo teamDataResponse = JsonConvert.DeserializeObject<TeamInfo>(teamDataJSON);
            return teamDataResponse.teams;
        }
    }
}
