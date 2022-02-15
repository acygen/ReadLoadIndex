﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.Data.SQLite;

namespace ReadLoadIndex
{
    public partial class Form1 : Form
    {
        Dictionary<int, string> unitNameDic;
        Dictionary<long, PlayerData> playerDatas = new Dictionary<long, PlayerData>();
        static SQLiteConnection cn;
        List<(int, string)> showUnitDataFromdb = new List<(int, string)>();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileName ofn = new OpenFileName();
            ofn.structSize = System.Runtime.InteropServices.Marshal.SizeOf(ofn);
            ofn.filter = "txt(*.txt)\0";  //指定打开格式
            ofn.file = new string(new char[256]);
            ofn.maxFile = ofn.file.Length;
            ofn.fileTitle = new string(new char[64]);
            ofn.maxFileTitle = ofn.fileTitle.Length;
            ofn.initialDir = Application.StartupPath;//默认路径
            ofn.title = "打开LoadIndex.txt";
            ofn.defExt = "txt";
            ofn.flags = 0x00080000 | 0x00001000 | 0x00000800 | 0x00000200 | 0x00000008;//OFN_EXPLORER|OFN_FILEMUSTEXIST|OFN_PATHMUSTEXIST| OFN_ALLOWMULTISELECT|OFN_NOCHANGEDIR
            //打开windows框
            if (DllTest.GetOpenFileName(ofn))
            {
                ImportPlayer(ofn.file);
            }
        }
        private void ImportPlayer(string path)
        {
            string txtRead = File.ReadAllText(path);
            LoadDataBody loadDataBody = JsonConvert.DeserializeObject<LoadDataBody>(txtRead);
            //richTextBox1.Clear();
            //PrintResult(loadDataBody);
            AddPlayer(loadDataBody);

        }
        private void PrintResult(LoadDataBody json)
        {
            if (unitNameDic == null)
            {
                loadNameDic();
            }
            Dictionary<int, int> unitLoveDic = new Dictionary<int, int>();
            foreach(UserChara chara in json.user_chara_info)
            {
                if (!unitLoveDic.ContainsKey(chara.chara_id * 100 + 1))
                {
                    unitLoveDic.Add(chara.chara_id * 100 + 1, chara.love_level);
                }
            }
            //string total = "";
            foreach(UnitDataS unit in json.unit_list)
            {
                int love = unitLoveDic.TryGetValue(unit.id, out var love0) ? love0 : 1;
                string equip = unit.GetEquipList();
                int[] skill = unit.GetSkillLevelInfo();
                string line = $"{unit.unit_level} {unit.unit_rarity} {love} {unit.promotion_level} {equip} {skill[0]}/{skill[1]}/{skill[2]}/{skill[3]} {unit.GetUek()}";
                string unitName = unitNameDic.TryGetValue(unit.id, out var value) ? value : "???";
                richTextBox1.AppendText($"UnitID:{unit.id}({unitName})\r\n");
                richTextBox1.AppendText(line + "\r\n");
            }
        }
        private void AddPlayer(LoadDataBody json)
        {
            if (unitNameDic == null)
            {
                loadNameDic();
                LoadSQLUnitData();
            }
            if (playerDatas.ContainsKey(json.user_info.viewer_id))
            {
                if(MessageBox.Show("已经有该玩家数据，是否覆盖？","提示",buttons: MessageBoxButtons.YesNo) == DialogResult.No)
                {
                    return;
                }
                playerDatas.Remove(json.user_info.viewer_id);
            }
            Dictionary<int, int> unitLoveDic = new Dictionary<int, int>();
            foreach (UserChara chara in json.user_chara_info)
            {
                if (!unitLoveDic.ContainsKey(chara.chara_id * 100 + 1))
                {
                    unitLoveDic.Add(chara.chara_id * 100 + 1, chara.love_level);
                }
            }
            //string total = "";
            Dictionary<int, string[]> unitDic2 = new Dictionary<int, string[]>();
            foreach (UnitDataS unit in json.unit_list)
            {
                int love = unitLoveDic.TryGetValue(unit.id, out var love0) ? love0 : 1;
                string equip = unit.GetEquipList();
                int[] skill = unit.GetSkillLevelInfo();
                List<string> data = new List<string>();
                string line = $"{unit.unit_level} {unit.unit_rarity} {love} {unit.promotion_level} {equip} {skill[0]}/{skill[1]}/{skill[2]}/{skill[3]} {unit.GetUek()}";
                data.Add(unit.unit_level.ToString());
                data.Add(unit.unit_rarity.ToString());
                data.Add(love.ToString());
                data.Add(unit.promotion_level.ToString());
                data.Add(equip);
                data.Add($"{ skill[0]}/{ skill[1]}/{ skill[2]}/{ skill[3]}");
                data.Add(unit.GetUek().ToString());

                unitDic2.Add(unit.id, data.ToArray());
            }
            PlayerData playerData = new PlayerData();
            playerData.name = System.Text.RegularExpressions.Regex.Unescape(json.user_info.user_name);
            playerData.boxDic = unitDic2;
            playerDatas.Add(json.user_info.viewer_id,playerData);
            richTextBox1.AppendText($"成功添加{playerData.name}的数据!\r\n");

        }
        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void loadNameDic()
        {
            unitNameDic = new Dictionary<int, string>();
            string path = Application.StartupPath + "/unitNameDic.json";
            if (File.Exists(path))
            {
                string dic = File.ReadAllText(path);
                unitNameDic = JsonConvert.DeserializeObject<Dictionary<int, string>>(dic);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            AllPlayerData allPlayerData = new AllPlayerData();
            Dictionary<int, string> avUnit = new Dictionary<int, string>();
            allPlayerData.playerDatas = new List<PlayerData>();

            foreach(var pair in playerDatas)
            {
                foreach(var p2 in pair.Value.boxDic)
                {
                    if (!avUnit.ContainsKey(p2.Key))
                    {
                        avUnit.Add(p2.Key, unitNameDic.TryGetValue(p2.Key, out var nn) ? nn : "???");
                    }
                }
                allPlayerData.playerDatas.Add(pair.Value);
            }
            allPlayerData.dbList = showUnitDataFromdb;
            if (showUnitDataFromdb.Count> 0)
            {
                avUnit.Clear();
                showUnitDataFromdb.ForEach(a => avUnit.Add(a.Item1, a.Item2));
            }
            allPlayerData.allUnitDic = avUnit;

            EXCELHelper.SaveExcel(allPlayerData);
            richTextBox1.AppendText($"成功导出EXCEL!\r\n");

        }
        private void DownloadDB()
        {
            try
            {
                var url = "https://lulubot.xyz/pcr/db_pcr.php";
                var save = Application.StartupPath + "/redive_cn.db";
                using (var web = new WebClient())
                {
                    web.DownloadFile(url, save);
                }
                MessageBox.Show("成功，请重启更新");
                this.Invoke(finish0);
            }
            catch(Exception ex)
            {
                MessageBox.Show("ERROR!" + ex.Message + "\n" + ex.StackTrace);
            }
        }
        delegate void finishDownload();
        finishDownload finish0;
        private void finishDownload_0()
        {
            button3.Enabled = true;
            button3.Text = "下载已经完成！";

        }
        private void button3_Click(object sender, EventArgs e)
        {
            finish0 = finishDownload_0;
            button3.Text = "正在下载...";
            button3.Enabled = false;
            Task a = Task.Run(DownloadDB);
            //DownloadDB();
        }
        private void LoadSQLUnitData()
        {
            try
            {
                showUnitDataFromdb.Clear();
                string path = Application.StartupPath + "/redive_cn.db";
                if (File.Exists(path))
                {
                    cn = new SQLiteConnection("data source=" + path);
                    if (cn.State != System.Data.ConnectionState.Open)
                    {
                        cn.Open();
                    }
                    string queryString = "SELECT * FROM " + "unit_data";
                    var dbCommand = cn.CreateCommand();
                    dbCommand.CommandText = queryString;
                    SQLiteDataReader reader = dbCommand.ExecuteReader();

                    while (reader.Read())
                    {
                        int id = reader.GetInt32(reader.GetOrdinal("unit_id"));
                        string name = reader.GetString(reader.GetOrdinal("unit_name"));
                        if (id <= 190000)
                        {
                            showUnitDataFromdb.Add((id, name));
                        }
                    }
                    cn.Close();
                }
                else
                {
                    MessageBox.Show("读取db失败！db不存在！" );

                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("读取db失败！" + ex.Message + ex.StackTrace);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string folderPath = "";
            folderBrowserDialog1.Description = "选择文件夹";
            folderBrowserDialog1.RootFolder = Environment.SpecialFolder.Desktop;
            folderBrowserDialog1.ShowNewFolderButton = true;
            if(folderBrowserDialog1.ShowDialog()== DialogResult.OK)
            {
                folderPath = folderBrowserDialog1.SelectedPath;
                string[] files = Directory.GetFiles(folderPath);
                if (files == null || files.Length <= 0)
                {
                    MessageBox.Show("空文件夹！");
                    return;
                }
                foreach(string path in files)
                {
                    if (Path.GetExtension(path).Contains("txt"))
                    {
                        ImportPlayer(path);
                    }
                }
            }
        }
    }
}