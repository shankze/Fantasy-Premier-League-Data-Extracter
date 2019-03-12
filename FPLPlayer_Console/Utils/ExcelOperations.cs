using FPLPlayer_Console.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FPLPlayer_Console.Utils
{
    class ExcelOperations
    {

        public static void PopulateResults(List<Element> playerList,List<History> statsList,List<Team> teamsList)
        {
            Dictionary<int, string> playerPosDict = CreatePlayerPositionDict();
            string OutputFilePath = @"C:\Users\sbekal\Dropbox\FPL\FPL_Output.xlsx";
            XSSFWorkbook workbook = GetWorkbook(OutputFilePath);
            ISheet outputSheet = workbook.GetSheet("Sheet1");
            int rowIndex = 1;
            int noOfPlayers = statsList.Count;
            int i = 1;

            WriteHeaders(outputSheet, 0, 0);

            foreach (History playerStat in statsList)
            {
                Element playerData = playerList.Find(item => item.id == playerStat.element);
                int columnStart = 0;
                //Player Info
                WriteToCell(outputSheet, playerStat.element.ToString(), rowIndex, columnStart + 0);
                Team playerTeam = teamsList.Find(item => item.id == playerData.team);
                WriteToCell(outputSheet, playerTeam.name, rowIndex, columnStart + 1);
                WriteToCell(outputSheet, playerData.first_name.ToString(), rowIndex, columnStart + 2);
                WriteToCell(outputSheet, playerData.second_name, rowIndex, columnStart + 3);
                columnStart++;
                string playerName = playerData.first_name + " " + playerData.second_name;
                WriteToCell(outputSheet, playerName, rowIndex, columnStart + 3);
                string position = "";
                playerPosDict.TryGetValue(playerData.element_type, out position);
                WriteToCell(outputSheet, position, rowIndex, columnStart + 4);
                double playerValue = playerStat.value / 10.0;
                WriteToCell(outputSheet, playerValue.ToString(), rowIndex, columnStart + 5);
                columnStart = columnStart + 2;
                //Game Info
                Team oppTeam = teamsList.Find(item => item.id == playerStat.opponent_team);
                WriteToCell(outputSheet, oppTeam.name, rowIndex, columnStart + 4);
                if (playerStat.was_home == true)
                {
                    WriteToCell(outputSheet, playerStat.team_h_score.ToString(), rowIndex, columnStart + 5);
                    WriteToCell(outputSheet, playerStat.team_a_score.ToString(), rowIndex, columnStart + 6);
                    WriteToCell(outputSheet, "Home", rowIndex, columnStart + 7);
                }
                else
                {
                    WriteToCell(outputSheet, playerStat.team_a_score.ToString(), rowIndex, columnStart + 5);
                    WriteToCell(outputSheet, playerStat.team_h_score.ToString(), rowIndex, columnStart + 6);
                    WriteToCell(outputSheet, "Away", rowIndex, columnStart + 7);
                }
                WriteToCell(outputSheet, playerStat.kickoff_time_formatted.ToString(), rowIndex, columnStart + 8);
                WriteToCell(outputSheet, playerStat.round.ToString(), rowIndex, columnStart + 9);
                WriteToCell(outputSheet, playerStat.minutes.ToString(), rowIndex, columnStart + 10);
                //Points
                WriteToCell(outputSheet, playerStat.total_points.ToString(), rowIndex, columnStart + 11);
                WriteToCell(outputSheet, playerStat.bonus.ToString(), rowIndex, columnStart + 12);
                WriteToCell(outputSheet, playerStat.bps.ToString(), rowIndex, columnStart + 13);
                //Index
                WriteToCell(outputSheet, playerStat.influence.ToString(), rowIndex, columnStart + 14);
                WriteToCell(outputSheet, playerStat.creativity.ToString(), rowIndex, columnStart + 15);
                WriteToCell(outputSheet, playerStat.threat.ToString(), rowIndex, columnStart + 16);
                WriteToCell(outputSheet, playerStat.ict_index.ToString(), rowIndex, columnStart + 17);
                WriteToCell(outputSheet, playerStat.ea_index.ToString(), rowIndex, columnStart + 18);
                //Attacking
                WriteToCell(outputSheet, playerStat.goals_scored.ToString(), rowIndex, columnStart + 19);
                WriteToCell(outputSheet, playerStat.assists.ToString(), rowIndex, columnStart + 20);
                WriteToCell(outputSheet, playerStat.open_play_crosses.ToString(), rowIndex, columnStart + 21);
                WriteToCell(outputSheet, playerStat.big_chances_created.ToString(), rowIndex, columnStart + 22);
                WriteToCell(outputSheet, playerStat.key_passes.ToString(), rowIndex, columnStart + 23);
                WriteToCell(outputSheet, playerStat.winning_goals.ToString(), rowIndex, columnStart + 24);
                WriteToCell(outputSheet, playerStat.attempted_passes.ToString(), rowIndex, columnStart + 25);
                WriteToCell(outputSheet, playerStat.completed_passes.ToString(), rowIndex, columnStart + 26);
                WriteToCell(outputSheet, playerStat.dribbles.ToString(), rowIndex, columnStart + 27);
                WriteToCell(outputSheet, playerStat.offside.ToString(), rowIndex, columnStart + 28);
                WriteToCell(outputSheet, playerStat.big_chances_missed.ToString(), rowIndex, columnStart + 29);
                WriteToCell(outputSheet, playerStat.target_missed.ToString(), rowIndex, columnStart + 30);
                //Defending
                WriteToCell(outputSheet, playerStat.clean_sheets.ToString(), rowIndex, columnStart + 31);
                WriteToCell(outputSheet, playerStat.goals_conceded.ToString(), rowIndex, columnStart + 32);
                WriteToCell(outputSheet, playerStat.own_goals.ToString(), rowIndex, columnStart + 33);
                WriteToCell(outputSheet, playerStat.penalties_saved.ToString(), rowIndex, columnStart + 34);
                WriteToCell(outputSheet, playerStat.penalties_missed.ToString(), rowIndex, columnStart + 35);
                WriteToCell(outputSheet, playerStat.saves.ToString(), rowIndex, columnStart + 36);
                WriteToCell(outputSheet, playerStat.clearances_blocks_interceptions.ToString(), rowIndex, columnStart + 37);
                WriteToCell(outputSheet, playerStat.recoveries.ToString(), rowIndex, columnStart + 38);
                WriteToCell(outputSheet, playerStat.tackles.ToString(), rowIndex, columnStart + 39);
                WriteToCell(outputSheet, playerStat.penalties_conceded.ToString(), rowIndex, columnStart + 40);
                WriteToCell(outputSheet, playerStat.errors_leading_to_goal.ToString(), rowIndex, columnStart + 41);
                WriteToCell(outputSheet, playerStat.errors_leading_to_goal_attempt.ToString(), rowIndex, columnStart + 42);
                WriteToCell(outputSheet, playerStat.tackles.ToString(), rowIndex, columnStart + 43);
                //Disciplinary
                WriteToCell(outputSheet, playerStat.yellow_cards.ToString(), rowIndex, columnStart + 44);
                WriteToCell(outputSheet, playerStat.red_cards.ToString(), rowIndex, columnStart + 45);
                WriteToCell(outputSheet, playerStat.fouls.ToString(), rowIndex, columnStart + 46);
                //Fantasy Related
                //WriteToCell(outputSheet, playerStat.value.ToString(), rowIndex, columnStart + 47);
                columnStart = columnStart - 1;
                WriteToCell(outputSheet, playerStat.transfers_balance.ToString(), rowIndex, columnStart + 48);
                WriteToCell(outputSheet, playerStat.transfers_balance.ToString(), rowIndex, columnStart + 49);
                WriteToCell(outputSheet, playerStat.selected.ToString(), rowIndex, columnStart + 50);
                WriteToCell(outputSheet, playerStat.transfers_in.ToString(), rowIndex, columnStart + 51);
                WriteToCell(outputSheet, playerStat.transfers_out.ToString(), rowIndex, columnStart + 52);


                rowIndex++;
                i++;
            }
            SaveToFilestream(workbook, OutputFilePath);
        }

        private static Dictionary<int, string> CreatePlayerPositionDict()
        {
            Dictionary<int, string> playerPosDict = new Dictionary<int, string>();
            playerPosDict.Add(1, "Goalkeeper");
            playerPosDict.Add(2, "Defender");
            playerPosDict.Add(3, "Midfielder");
            playerPosDict.Add(4, "Forward");
            return playerPosDict;
        }

        private static void WriteHeaders(ISheet outputSheet, int rowIndex, int columnStart)
        {
            WriteToCell(outputSheet, "Player ID", rowIndex, columnStart + 0); //0 columnStart=0
            WriteToCell(outputSheet, "Team", rowIndex, columnStart + 1); //1 columnStart=0
            WriteToCell(outputSheet, "First Name", rowIndex, columnStart + 2); //2 columnStart=0
            WriteToCell(outputSheet, "Last Name", rowIndex, columnStart + 3); //3 columnStart=0
            columnStart++; 
            WriteToCell(outputSheet, "Player Name", rowIndex, columnStart + 3); //4 columnStart=1
            WriteToCell(outputSheet, "Position", rowIndex, columnStart + 4); //5 columnStart=1
            WriteToCell(outputSheet, "value", rowIndex, columnStart + 5); //6 columnStart=1
            columnStart = columnStart + 2;
            //Game Info
            WriteToCell(outputSheet, "opponent_team", rowIndex, columnStart + 4); //7 columnStart=3
            WriteToCell(outputSheet, "For Score", rowIndex, columnStart + 5); //8 columnStart=3
            WriteToCell(outputSheet, "Against Score", rowIndex, columnStart + 6); //9 columnStart=3
            WriteToCell(outputSheet, "Home/Away", rowIndex, columnStart + 7);
            WriteToCell(outputSheet, "kickoff_time_formatted", rowIndex, columnStart + 8);
            WriteToCell(outputSheet, "round", rowIndex, columnStart + 9);
            WriteToCell(outputSheet, "minutes", rowIndex, columnStart + 10);
            //Points
            WriteToCell(outputSheet, "total_points", rowIndex, columnStart + 11);
            WriteToCell(outputSheet, "bonus", rowIndex, columnStart + 12);
            WriteToCell(outputSheet, "bps", rowIndex, columnStart + 13);
            //Index
            WriteToCell(outputSheet, "influence", rowIndex, columnStart + 14);
            WriteToCell(outputSheet, "creativity", rowIndex, columnStart + 15);
            WriteToCell(outputSheet, "threat", rowIndex, columnStart + 16);
            WriteToCell(outputSheet, "ict_index", rowIndex, columnStart + 17);
            WriteToCell(outputSheet, "ea_index", rowIndex, columnStart + 18);
            //Attacking
            WriteToCell(outputSheet, "goals_scored", rowIndex, columnStart + 19);
            WriteToCell(outputSheet, "assists", rowIndex, columnStart + 20);
            WriteToCell(outputSheet, "open_play_crosses", rowIndex, columnStart + 21);
            WriteToCell(outputSheet, "big_chances_created", rowIndex, columnStart + 22);
            WriteToCell(outputSheet, "key_passes", rowIndex, columnStart + 23);
            WriteToCell(outputSheet, "winning_goals", rowIndex, columnStart + 24);
            WriteToCell(outputSheet, "attempted_passes", rowIndex, columnStart + 25);
            WriteToCell(outputSheet, "completed_passes", rowIndex, columnStart + 26);
            WriteToCell(outputSheet, "dribbles", rowIndex, columnStart + 27);
            WriteToCell(outputSheet, "offside", rowIndex, columnStart + 28);
            WriteToCell(outputSheet, "big_chances_missed", rowIndex, columnStart + 29);
            WriteToCell(outputSheet, "target_missed", rowIndex, columnStart + 30);
            //Defending
            WriteToCell(outputSheet, "clean_sheets", rowIndex, columnStart + 31);
            WriteToCell(outputSheet, "goals_conceded", rowIndex, columnStart + 32);
            WriteToCell(outputSheet, "own_goals", rowIndex, columnStart + 33);
            WriteToCell(outputSheet, "penalties_saved", rowIndex, columnStart + 34);
            WriteToCell(outputSheet, "penalties_missed", rowIndex, columnStart + 35);
            WriteToCell(outputSheet, "saves", rowIndex, columnStart + 36);
            WriteToCell(outputSheet, "clearances_blocks_interceptions", rowIndex, columnStart + 37);
            WriteToCell(outputSheet, "recoveries", rowIndex, columnStart + 38);
            WriteToCell(outputSheet, "tackles", rowIndex, columnStart + 39);
            WriteToCell(outputSheet, "penalties_conceded", rowIndex, columnStart + 40);
            WriteToCell(outputSheet, "errors_leading_to_goal", rowIndex, columnStart + 41);
            WriteToCell(outputSheet, "errors_leading_to_goal_attempt", rowIndex, columnStart + 42);
            WriteToCell(outputSheet, "tackles", rowIndex, columnStart + 43);
            //Disciplinary
            WriteToCell(outputSheet, "yellow_cards", rowIndex, columnStart + 44);
            WriteToCell(outputSheet, "red_cards", rowIndex, columnStart + 45);
            WriteToCell(outputSheet, "fouls", rowIndex, columnStart + 46); //49 columnStart=3
            //Fantasy Related
            //WriteToCell(outputSheet, "value", rowIndex, columnStart + 47);
            columnStart = columnStart - 1;
            WriteToCell(outputSheet, "transfers_balance", rowIndex, columnStart + 48); //50 columnStart=2
            WriteToCell(outputSheet, "transfers_balance", rowIndex, columnStart + 49);
            WriteToCell(outputSheet, "selected", rowIndex, columnStart + 50);
            WriteToCell(outputSheet, "transfers_in", rowIndex, columnStart + 51);
            WriteToCell(outputSheet, "transfers_out", rowIndex, columnStart + 52);
        }

        private static void SaveToFilestream(XSSFWorkbook workbook, string path)
        {
            using (FileStream sw = File.Create(path))
            {
                workbook.Write(sw);
            }
        }

        private static XSSFWorkbook GetWorkbook(string path)
        {
            XSSFWorkbook workbook;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(file);
            }
            return workbook;
        }

        private static void WriteToCell(ISheet sheet, string value, int rowNum, int columnNum)
        {
            if (value != null)
            {
                try
                {
                    if (sheet.GetRow(rowNum) == null)
                    {
                        sheet.CreateRow(rowNum);
                    }
                    if (sheet.GetRow(rowNum).GetCell(columnNum) == null)
                    {
                        sheet.GetRow(rowNum).CreateCell(columnNum);
                    }
                    sheet.GetRow(rowNum).GetCell(columnNum).SetCellValue(value);
                }
                catch (Exception ex)
                {
                }
            }
        }
    }
}
