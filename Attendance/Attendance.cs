/*
    Program Name : Attendance Program
    Creation Date : 2018/03/22
    Creation Reason : 영어 출석부가 필요하다고 생각해서 영어 선생님과 같이 사용할 수 있는 출석부 개발
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.VisualBasic;

namespace Attendance
{
    public partial class Attendance : Form
    {
        // 각 date, count 값 * 와함께 변수화 시켜서 저장
        private static DateTimePicker[] dateArr = null;
        private static NumericUpDown[] updownArr = null;
        private static Button[] btnArr = null;


        // 폴더 변수 전역으로
        string mainFolderDir = "c:\\SEIAttendanceBook";

        public Attendance()
        {

            // 폴더 있는지 체크 하고 없으면 만들어 줌
            DirectoryInfo folder = new DirectoryInfo(mainFolderDir);
            if (!folder.Exists)
            {
                folder.Create();
            }

            InitializeComponent();

            // static 변수로 담고 있기
            GetDateTimePickerList();
            GetNumericUpDownList();
            GetButtonList();
            
            // member grid 채우기
            DisplayGridMember();

        }

        // 신규 멤버 추가하는 메소드
        private void AddMemBtn_Click(object sender, EventArgs e)
        {
            // 추가할 멤버 이름 입력
            string newMember = Interaction.InputBox("Please Input New Member.", "New Member", "jyh(1987)", 10, 10);
            // 취소해버리니깐 값이 안넘어간다 IsNullOrWhiteSpace 로 처리
            // 취소 했는지 체크 하고 return 
            if (String.IsNullOrWhiteSpace(newMember)) 
            {
                return;
            }

            // 추가할 멤버 폴더 생성
            string memberFolder = mainFolderDir + "\\" + newMember;
            DirectoryInfo folder = new DirectoryInfo(memberFolder);
            if (!folder.Exists)
            {
                // 폴더 생성
                folder.Create();

                // 초기 출석부 파일 생성
                string filePath = memberFolder  + "\\" + "0001.txt";
                FileInfo file = new FileInfo(filePath);
                if (!file.Exists)
                {
                    FileStream newFile = file.Create();
                    // 파일 안닫아주면 오류 발생
                    newFile.Close();
                }

                DataGridViewRow row = new DataGridViewRow();
                // 신규 멤버 그리드에 추가
                AddGridMember(newMember);

            }
            else
            {
                // 이미 있는 member
                MessageBox.Show("Aleady Added Member.");
            }
        }

        // 멤버들 그리드에 디스플레이 하는 메소드
        private void DisplayGridMember()
        {
            // 쌓이는거 막기 위해 clear
            GridMemList.Rows.Clear();
            
            DirectoryInfo[] folderList = GetMemberFolderList();

            for (int i=0; i<folderList.Length; ++i)
            {
                AddGridMember(folderList[i].Name);
            }

        }

        // SEI 메인 폴더에 생성된 폴더 list 얻는 메소드(멤버리스트와 동일)
        private DirectoryInfo[] GetMemberFolderList()
        {
            DirectoryInfo folder = new DirectoryInfo(mainFolderDir);
            DirectoryInfo[] folderList = folder.GetDirectories();
            return folderList;
        }

        // 디스플레이 하기 위해서 멤버를 그리드에 넣는 실제 로직
        private void AddGridMember(string member)
        {
            DataGridViewRow row = new DataGridViewRow();
            row.CreateCells(GridMemList);

            row.Cells[0].Value = member;
            row.Cells[1].Value = GetCount(member); 

            GridMemList.Rows.Add(row);
        }

        // grid member에 들어가는 count 계산해주는 메소드 
        // 해당 멤버의 최신파일 기준으로 카운트 구한다.
        private string GetCount(string member)
        {
            // member folder 로 이동
            string folderPath = mainFolderDir + "\\" + member;
            DirectoryInfo dir = new DirectoryInfo(folderPath);
            FileInfo[] files = dir.GetFiles();

            // 최신파일 이름
            string latestFile = files[files.Length - 1].Name;

            string filePath = folderPath + "\\" + latestFile;
            string[] lines =  System.IO.File.ReadAllLines(@filePath);
            
            // 빈칸 빼고 카운트 올릴 것
            int count = 0;
            for (int i=0; i<lines.Length; ++i)
            {
                if (!String.IsNullOrEmpty(lines[i]))
                {
                    count += int.Parse(lines[i].Split('*')[1]);
                }
            }

            return count.ToString();
        }

        // SearchTextBar에 입력한 값 기준으로 Grid 필터링
        private void NameText_KeyDown(object sender, KeyEventArgs e)
        {
            // clear
            GridMemList.Rows.Clear();

            DirectoryInfo[] folderList = GetMemberFolderList();

            // 입력한 값이랑 일치하는지 for문 안에서 확인
            for (int i = 0; i < folderList.Length; ++i)
            {
                if (folderList[i].Name.Contains(NameText.Text))
                {
                    AddGridMember(folderList[i].Name);
                }
            }
        }

        // grid에서 선택한 멤버 삭제하는 메소드
        // 폴더와, 폴더안의 파일들을 삭제한다.
        private void DelMemBtn_Click(object sender, EventArgs e)
        {
            // 신규 생성시 선택 안되어 있으면 실행 안되게
            if (String.IsNullOrEmpty(SelectedMemberName.Text))
            {
                MessageBox.Show("Please Select Member!");
                return;
            }

            // cell 에서 선택된 멤버 이름
            DataGridViewSelectedCellCollection selectedMember = GridMemList.SelectedCells;
            string count = selectedMember.Count.ToString();

            // 몇명 삭제할건지 물어봄
            DialogResult dialogResult = 
            MessageBox.Show("Do you want to delete selected " + count  + " members?", 
                            "Delete Member?", 
                            MessageBoxButtons.YesNo);

            if (dialogResult.ToString() == "Yes")
            {
                for (int i = 0; i < selectedMember.Count; ++i)
                {
                    // 삭제할 멤버
                    string selectedMem = selectedMember[i].Value.ToString();
                    // 삭제할 멤버 폴더
                    string memberFolder = mainFolderDir + "\\" + selectedMem;
                  
                    DirectoryInfo folder = new DirectoryInfo(memberFolder);

                    // 폴더내부의 파일들 먼저 삭제
                    DeletFilesInFolder(folder);

                    // 폴더 삭제
                    folder.Delete();
                }

                // refresh
                DisplayGridMember();

                // 검색 값 초기화
                NameText.Text = "";

            }

        }

        // 폴더 내부의 파일들 삭제 하기 위함
        private void DeletFilesInFolder(DirectoryInfo folder)
        {
            // 초기 출석부 파일 생성
            FileInfo [] files = folder.GetFiles();
            for (int i=0; i<files.Length; ++i)
            {
                files[i].Delete();
            }

        }

        // 멤버 이름 변경을 위한 메소드
        private void RevMemBtn_Click(object sender, EventArgs e)
        {
            // 신규 생성시 선택 안되어 있으면 실행 안되게
            if (String.IsNullOrEmpty(SelectedMemberName.Text))
            {
                MessageBox.Show("Please Select Member!");
                return;
            }

            // 선택된 멤버 이름
            string prevName = GridMemList.SelectedCells[0].Value.ToString();

            // 변경할 이름
            string revisedName = 
            Interaction.InputBox(   "Please Input Revised Name \n'" + prevName + "' -> ", 
                                    "Revise Name of Member",
                                    prevName, 
                                    10, 
                                    10);

            // 같거나 값이 없으면 취소 하는 효과
            if (revisedName == prevName || String.IsNullOrWhiteSpace(revisedName)) 
            { 
                return;
            }

            // 멤버 폴더 이름
            string memberFolder = mainFolderDir + "\\" + revisedName;

            // 폴더 유무 체크
            DirectoryInfo folder = new DirectoryInfo(memberFolder);
            if (!folder.Exists)
            {

                // 폴더 네임 변경
                string prevFolderPath = mainFolderDir + "\\" + prevName;
                Directory.Move(@prevFolderPath, @memberFolder);

                // grid Name 변경
                GridMemList.SelectedCells[0].Value = revisedName;

                // selected Label 값 변경
                SelectedMemberName.Text = revisedName;

            }
            else
            {
                MessageBox.Show("Aleady Added Member.");
            }

        }

        // tab 할때 에러 발생해서 key값 받아서 false 로 변경해줌
        // 여러경우 시도 했는데 이렇게 밖에 안되네
        bool tabCheck = true;

        // 셀 선택 변경시 작동
        private void GridMemList_SelectionChanged(object sender, EventArgs e)
        {
            if (!tabCheck)
            {
                tabCheck = true;
                return;
            }
            try
            {
                // counts 컬럼 선택했으면 member column으로 옮겨줄 것
                if (GridMemList.SelectedCells[0].OwningColumn.Index.ToString() == "1")
                {
                    // row index 구함
                    int rowIndex = GridMemList.SelectedCells[0].OwningRow.Index;
                    // current cell 변경 ( counts column 선택해도 자동으로 memberlist로 변경함 )
                    GridMemList.CurrentCell = GridMemList.Rows[rowIndex].Cells[0];
                    return; // return 안해주면 아래 실제 로직을 한번 더 타는데 이래서 두번 출력된다.
                }
            }
            catch (Exception ee)
            {
                // 더이상 삭제 할게 없으면 실행되는 로직

                // 검색, 삭제 할때 인덱스 에러 발생
                // 위젯 모두 false 변경해주고
                ChangeVisibleFalse(dateArr, updownArr, btnArr);
                
                // selectedmembername 값 초기화
                SelectedMemberName.Text = "";

                // 실제 로직처리 안되게끔 리턴
                return;
            }

            // 실제 로직 처리
            SelectCell();
        }

        // tab 키 누르면 자동으로 count column으로 이동되는 오류 막음
        private void GridMemList_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode.ToString() == "Tab")
            {
                tabCheck = false;
            }
        }


        // 실제 연산 로직 여기서 모두 처리
        private void SelectCell()
        {
            string selectCellValue = "";

            // 선택된 cell 값
            selectCellValue = GridMemList.SelectedCells[0].Value.ToString();

            // 현재 선택하고 있는 member 표시해줌
            SelectedMemberName.Text = selectCellValue;

            // History 부분 디스플레이 Logic
            string folderPath = mainFolderDir + "\\" + selectCellValue;

            // 최신파일 선택(사람바뀔때)
            SelectLatestFile(folderPath);

        }

        // 히스토리에서 최신파일 선택 로직(사람바뀔때)
        private void SelectLatestFile (string memberFolder)
        {
            // 쌓이는 것 방지하기 위해서 clear()
            HistoryCombo.Items.Clear();

            DirectoryInfo folder = new DirectoryInfo(memberFolder);
            FileInfo[] files = folder.GetFiles();
            for (int i = 0; i < files.Length; ++i)
            {
                HistoryCombo.Items.Add(files[i].Name);
            }

            // 출석파일 중 최신 파일 찾기
            int hisCount = HistoryCombo.Items.Count;

            // 처음 시작 최신 파일로 출발 
            // history 부분 선택
            HistoryCombo.SelectedIndex = hisCount - 1;

        }

        // 히스토리 기반으로 출석부 실제 내용 출력
        private void DisplayHistory(string filePath)
        {

            //시간 출력 test
            //System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            //sw.Reset();
            //sw.Start();

            // visible 모두 false 로 해주기, 수업 아무것도 안들은 사람은 이전 사람 기록 그대로 보여진다.
            ChangeVisibleFalse(dateArr, updownArr, btnArr);

            // string text = System.IO.File.ReadAllText(@filePath);
            string [] lines = System.IO.File.ReadAllLines(@filePath);


            int count = 0; // 선택한 파일로 grid count revise 할 count
            for (int i=0; i<lines.Length; ++i)
            {
                // 날짜 불러오기
                // dateArr[i].Visible = true;
                dateArr[i].Show();
                string[] split = lines[i].Split('*');
                DateTime convertedDate;
                convertedDate = Convert.ToDateTime(split[0]);
                dateArr[i].Value = convertedDate;

                // 횟수 불러오기
                updownArr[i].Value = int.Parse(split[1]);
                // updownArr[i].Visible = true;
                updownArr[i].Show();
                
                // 체크 버튼 불러오기 
                // btnArr[i].Visible = true;
                btnArr[i].Show();
                btnArr[i].BackColor = System.Drawing.Color.Lime;

                // 선택한 파일로 grid count revise 할 count++
                count += int.Parse(split[1]);
            }

            // row index 값 구함
            int rowIndex = GridMemList.SelectedCells[0].OwningRow.Index;

            // 값 변경, 현재 선택한 파일의 최대 count로 
            GridMemList.Rows[rowIndex].Cells[1].Value = count.ToString();

            //시간 출력 결과
            //sw.Stop();
            //MessageBox.Show((sw.ElapsedMilliseconds/1000.0f).ToString());
        }

        // NewDate 했을때 동작하는 메소드
        private void NewDateBtn_Click(object sender, EventArgs e)
        {
            // 신규 생성시 선택 안되어 있으면 실행 안되게
            if (String.IsNullOrEmpty(SelectedMemberName.Text))
            {
                MessageBox.Show("Please Select Member!");
                return;
            }

            for (int i=0; i<dateArr.Length; ++i)
            {
                // dateArr[i].Visible 가 false 면 
                // 여기서 부터 버튼 추가 
                if (!dateArr[i].Visible)
                {
                    // 체크 과정 하나 더 필요함 기록 안된게 두번이상 못 들어가게
                    // 이전 버튼의 색깔 확인해서 break;
                    if (i > 0 && btnArr[i - 1].BackColor == System.Drawing.Color.Red)
                    {
                        MessageBox.Show("Don't hurry up, please save it one day.");
                        break;
                    }

                    //dateArr[i].Visible = true;
                    //updownArr[i].Visible = true;
                    //btnArr[i].Visible = true;

                    dateArr[i].Show();
                    updownArr[i].Show();
                    btnArr[i].Show();

                    // 버튼만 백그라운드 red 로 변경
                    btnArr[i].BackColor = System.Drawing.Color.Red;
                    break;
                }
            }
        }

        // 히스토리에서 선택한 대로 출석부 내용 출력
        private void HistoryCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            // select member 구하기
            string selectMem = SelectedMemberName.Text;
            
            // history 가 선택한 파일명
            string selectedFile = HistoryCombo.SelectedItem.ToString();

            // 히스토리에서 선택된 파일 패스 구하기
            string filePath = mainFolderDir + "\\" + selectMem + "\\" + selectedFile;

            // 사람과 history 기반으로 출석부 내용 출력
            DisplayHistory(filePath);
        }

        // visible 모두 false 로 변경해주는 것
        // history 값 변경할때 안없어지고 남아 있는 문제 있기 때문
        private void ChangeVisibleFalse(DateTimePicker[] dateArr, NumericUpDown[] updownArr, Button[] btnArr)
        {
            for (int i=0; i<dateArr.Length; ++i)
            {
                // dateArr[i].Visible = false;
                dateArr[i].Hide();
                // false 만 하니깐 다른값 조회했다가 풀면 이상한 날짜로 되어있다. 오늘 날짜로 초기화 진행
                dateArr[i].Value = DateTime.Now.Date;

                // updownArr[i].Visible = false;
                updownArr[i].Hide();
                updownArr[i].Value = 1;

                // btnArr[i].Visible = false;
                btnArr[i].Hide();
                btnArr[i].BackColor = System.Drawing.Color.Lime;
            }
        }

        // save logic
        private void SaveBtn_Click(object sender, EventArgs e)
        {
            // 신규 생성시 선택 안되어 있으면 실행 안되게
            if (String.IsNullOrEmpty(SelectedMemberName.Text))
            {
                MessageBox.Show("Please Select Member!");
                return;
            }

            // 저장 할건지 물어봄
            DialogResult dialogResult =
            MessageBox.Show("Do you want to Save?",
                            "Save?",
                            MessageBoxButtons.YesNo);

            if (dialogResult.ToString() == "Yes")
            {
                // select member 구하기
                string selectMem = SelectedMemberName.Text;

                // history 가 선택한 파일명
                string selectedFile = HistoryCombo.SelectedItem.ToString();

                // 히스토리에서 선택된 파일 패스 구하기
                string filePath = mainFolderDir + "\\" + selectMem + "\\" + selectedFile;

                // 임시 string array
                string[] tmpStrArr = new string[25];
                int indexCount = 0; // 실제 어레이 갯수 구하기 위해
                int totalCount = 0; // 실제 count 총량 구하기 위해서, 2,3 번 하루에 여러번 할수도 있으니
                for (int i=0; i<dateArr.Length; ++i)
                {
                    // 보여야지 저장할수 있게
                    if (dateArr[i].Visible)
                    {
                        string year = dateArr[i].Value.Date.Year.ToString();
                        string month = dateArr[i].Value.Date.Month.ToString();
                        string day = dateArr[i].Value.Date.Day.ToString();
                        string date = year + "/" + month + "/" + day;

                        string count = updownArr[i].Value.ToString();
                        tmpStrArr[indexCount] = date + "*" + count;
                        indexCount++; // 실제 어레이 갯수 구하기 위해
                        totalCount += int.Parse(updownArr[i].Value.ToString());
                    }

                }
                
                // 실제 파일에 쓸 array 에 값 전달해준다.
                string[] lines = new string[indexCount];
                for (int i=0; i<lines.Length; ++i)
                {
                    lines[i] = tmpStrArr[i];
                }

                // 파일에 쓰기 
                System.IO.File.WriteAllLines(@filePath, lines);

                // 버튼 색깔 바꿔주기 
                btnArr[lines.Length-1].BackColor = System.Drawing.Color.Lime;

                // count 변경 해주기 
                int gridIndex = GridMemList.SelectedCells[0].OwningRow.Index;
                GridMemList.Rows[gridIndex].Cells[1].Value = totalCount;
                
            }
        }

        // new page btn 작동 메소드, 
        // 파일명 0000.txt 맞춰서 생성
        private void NewPageBtn_Click(object sender, EventArgs e)
        {
            // 신규 생성시 선택 안되어 있으면 실행 안되게
            if (String.IsNullOrEmpty(SelectedMemberName.Text))
            {
                MessageBox.Show("Please Select Member!");
                return;
            }
            
            // 페이지 추가 할건지 물어봄
            DialogResult dialogResult =
            MessageBox.Show("Do you want to add new page?",
                            "Add New Page?",
                            MessageBoxButtons.YesNo);

            // yes 아니면 리턴해서 종료
            if (dialogResult.ToString() != "Yes")
                return;
            
            // 선택된 멤버 가져오기
            string member = SelectedMemberName.Text;

            // 멤버 폴더 패스 구하기
            string memberFolderPath = mainFolderDir + "\\" + member;

            // 폴더내의 파일들 가져오기
            DirectoryInfo dir = new DirectoryInfo(memberFolderPath);
            FileInfo [] files = dir.GetFiles();
            // 마지막 파일 가져오기
            string lastFileName = files[files.Length - 1].Name;
            // .txt 떼주기
            lastFileName = lastFileName.Replace(".txt", "");
            // +1해서 최신 파일 이름 만듬
            int newFileNum = int.Parse(lastFileName) + 1;
            // 숫자니깐 0000.txt 형태로 변경해줌
            string newFileName = newFileNum.ToString();
            int forCount = 4 - newFileName.Length;
            for (int i=0; i<forCount; ++i)
            {
                newFileName = "0" + newFileName;
            }
            // 파일 네임 완성
            newFileName = newFileName + ".txt";
            string filePath = memberFolderPath + "\\" + newFileName;
            FileInfo newFile = new FileInfo(filePath);
            // 파일 생성
            FileStream fs = newFile.Create();
            // 안닫아주면 에러 발생
            fs.Close();

            // history 추가, 선택 변경
            HistoryCombo.Items.Add(newFileName);
            
            // 출석파일 중 최신 파일 찾기
            int hisCount = HistoryCombo.Items.Count;

            // 처음 시작 최신 파일로 출발 
            // history 부분 선택
            HistoryCombo.SelectedIndex = hisCount - 1;

        }

        // help pdf 띄울 메소드
        private void HelpBtn_Click(object sender, EventArgs e)
        {
            // help 파일 볼건지 물어봄
            DialogResult dialogResult =
            MessageBox.Show("Do you want to see help page?",
                            "See Help Page?",
                            MessageBoxButtons.YesNo);

            // yes 아니면 리턴해서 종료
            if (dialogResult.ToString() != "Yes")
                return;
           
            string url = "https://github.com/fuding87/etc/raw/master/etcPrivate/WebContent/AttendanceBook/AttendanceBookHelp.pptx";
            System.Diagnostics.Process.Start(url);
        }

        // 출석부 디스플레이시 쉽게하기 위해서 array로 묶어줌
        private void GetDateTimePickerList()
        {
            dateArr = new DateTimePicker[25];
            for (int i=0; i<25; ++i)
            {
                dateArr[i] = Controls.Find("dateTimePicker" + (i + 1), true)[0] as DateTimePicker;
            }

        }
        // 출석부 디스플레이시 쉽게하기 위해서 array로 묶어줌
        private void GetNumericUpDownList()
        {
            updownArr = new NumericUpDown[25];
            for (int i = 0; i < 25; ++i)
            {
                updownArr[i] = Controls.Find("numericUpDown" + (i + 1), true)[0] as NumericUpDown;
            }
        }
        // 출석부 디스플레이시 쉽게하기 위해서 array로 묶어줌
        private void GetButtonList()
        {
            btnArr = new Button[25];
            for (int i = 0; i < 25; ++i)
            {
                btnArr[i] = Controls.Find("button" + (i + 1), true)[0] as Button;
            }

        }

    }
}
