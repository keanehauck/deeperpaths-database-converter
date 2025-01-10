// DeeperPathsConverter.cpp : Defines the entry point for the application.
//

#include "framework.h"
#include "DeeperPathsConverter.h"
#include <fstream>
#include <string>
#include <map>
#include <vector>
#include <regex>
#include <nlohmann/json.hpp>
#include "xlsxwriter.h"
#include <windows.h>
#include <shobjidl.h> // For IFileOpenDialog
#include <string>
#include <commdlg.h>  // Common dialog box

#include <nlohmann/json.hpp>
#include <iostream>
#include <string>

#pragma comment(lib,"user32.lib") 
#pragma comment(lib,"Ole32.lib") 


std::wstring OpenFileDialog() {
    // COM initialization
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    std::wstring filePath;

    if (SUCCEEDED(hr)) {
        // Create the File Open Dialog object
        IFileOpenDialog* pFileOpen;

        hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_ALL, IID_IFileOpenDialog, reinterpret_cast<void**>(&pFileOpen));

        if (SUCCEEDED(hr)) {
            // Set file type filters (only JSON files)
            COMDLG_FILTERSPEC filters[] = {
                { L"JSON Files", L"*.json" },
                { L"All Files", L"*.*" }
            };
            pFileOpen->SetFileTypes(2, filters);

            // Show the Open dialog box
            hr = pFileOpen->Show(NULL);

            // Process selected file
            if (SUCCEEDED(hr)) {
                IShellItem* pItem;
                hr = pFileOpen->GetResult(&pItem);
                if (SUCCEEDED(hr)) {
                    PWSTR pszFilePath;
                    hr = pItem->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath);

                    // Copy the selected file path
                    if (SUCCEEDED(hr)) {
                        filePath = pszFilePath;
                        CoTaskMemFree(pszFilePath);
                    }
                    pItem->Release();
                }
            }
            pFileOpen->Release();
        }
        CoUninitialize();
    }

    // Return the selected file path
    return filePath;
}

// Safe getter function that converts JSON values to appropriate types
template <typename T>
T safe_get(const nlohmann::json& j, const std::string& key, const T& default_value = T()) {
    try {
        if (j.contains(key)) {
            return j[key].get<T>();
        }
    }
    catch (const std::exception& e) {
        std::cerr << "Error retrieving key " << key << ": " << e.what() << std::endl;
    }
    return default_value;  // Return default value if key doesn't exist or an error occurs
}

std::string safe_get_string(const nlohmann::json& j, const std::string& key) {
    return safe_get<std::string>(j, key, "");
}

int safe_get_int(const nlohmann::json& j, const std::string& key) {
    return safe_get<int>(j, key, 0);
}

double safe_get_double(const nlohmann::json& j, const std::string& key) {
    return safe_get<double>(j, key, 0.0);
}

bool safe_get_bool(const nlohmann::json& j, const std::string& key) {
    return safe_get<bool>(j, key, false);
}


using json = nlohmann::json;

void createExcel(const nlohmann::json& data) {
    lxw_workbook* workbook = workbook_new("output.xlsx");
    lxw_worksheet* worksheet = workbook_add_worksheet(workbook, NULL);

    // Add headers
    worksheet_write_string(worksheet, 0, 0, "Participant ID", NULL);
    worksheet_write_string(worksheet, 0, 1, "Date", NULL);
    worksheet_write_string(worksheet, 0, 2, "Time", NULL);

    // Add settings columns
    worksheet_write_string(worksheet, 0, 3, "ChangeSuit", NULL);
    worksheet_write_string(worksheet, 0, 4, "GameOfScoreChange", NULL);
    worksheet_write_string(worksheet, 0, 5, "IncludeTimer", NULL);
    worksheet_write_string(worksheet, 0, 6, "MaxTimePerRound", NULL);
    worksheet_write_string(worksheet, 0, 7, "NumGames", NULL);
    worksheet_write_string(worksheet, 0, 8, "RoundOfScoreChange", NULL);

    int startCol = 9;
    int indices[10];

    for (int games = 1; games <= 10; ++games) {
        for (int rounds = 1; rounds <= 10; ++rounds) {
            for (int moves = 1; moves <= 8; ++moves) {
                worksheet_write_string(worksheet, 0, startCol++,
                    ("Game_" + std::to_string(games) + " Round_" + std::to_string(rounds) + " Move_" + std::to_string(moves) + " moveNumber").c_str(), NULL);
                worksheet_write_string(worksheet, 0, startCol++,
                    ("Game_" + std::to_string(games) + " Round_" + std::to_string(rounds) + " Move_" + std::to_string(moves) + " points").c_str(), NULL);
                worksheet_write_string(worksheet, 0, startCol++,
                    ("Game_" + std::to_string(games) + " Round_" + std::to_string(rounds) + " Move_" + std::to_string(moves) + " time").c_str(), NULL);
                worksheet_write_string(worksheet, 0, startCol++,
                    ("Game_" + std::to_string(games) + " Round_" + std::to_string(rounds) + " Move_" + std::to_string(moves) + " type").c_str(), NULL);
                worksheet_write_string(worksheet, 0, startCol++,
                    ("Game_" + std::to_string(games) + " Round_" + std::to_string(rounds) + " Move_" + std::to_string(moves) + " row").c_str(), NULL);
                worksheet_write_string(worksheet, 0, startCol++,
                    ("Game_" + std::to_string(games) + " Round_" + std::to_string(rounds) + " Move_" + std::to_string(moves) + " col").c_str(), NULL);
                worksheet_write_string(worksheet, 0, startCol++,
                    ("Game_" + std::to_string(games) + " Round_" + std::to_string(rounds) + " Move_" + std::to_string(moves) + " rank").c_str(), NULL);
                worksheet_write_string(worksheet, 0, startCol++,
                    ("Game_" + std::to_string(games) + " Round_" + std::to_string(rounds) + " Move_" + std::to_string(moves) + " suit").c_str(), NULL);
            }
            
            worksheet_write_string(worksheet, 0, startCol++,
                ("Game_" + std::to_string(games) + " Round_" + std::to_string(rounds) + " RoundScore").c_str(), NULL);
            worksheet_write_string(worksheet, 0, startCol++,
                ("Game_" + std::to_string(games) + " Round_" + std::to_string(rounds) + " Exploitation").c_str(), NULL);
            worksheet_write_string(worksheet, 0, startCol++,
                ("Game_" + std::to_string(games) + " Round_" + std::to_string(rounds) + " Exploration").c_str(), NULL);
            worksheet_write_string(worksheet, 0, startCol++,
                ("Game_" + std::to_string(games) + " Round_" + std::to_string(rounds) + " PercentUnexplored").c_str(), NULL);
            worksheet_write_string(worksheet, 0, startCol++,
                ("Game_" + std::to_string(games) + " Round_" + std::to_string(rounds) + " AvgRoundScore").c_str(), NULL);



        }

        worksheet_write_string(worksheet, 0, startCol++,
            ("Game_" + std::to_string(games) + " FinalScore").c_str(), NULL);

    }

    int row = 1;
    int col = 0;

    for (auto& participant : data["participants"].items()) {
        auto participant_data = participant.value();
        std::string participant_id = participant.key();
        std::string date = safe_get_string(participant_data["datetime"], "date");
        std::string time = safe_get_string(participant_data["datetime"], "time");
        auto settings = participant_data["settings"];
        int num_games = safe_get_int(settings, "numGames");


        col = 0; //reset column index

        // Write participant info (ID, Date, Time)
        worksheet_write_string(worksheet, row, col++, participant_id.c_str(), NULL);
        worksheet_write_string(worksheet, row, col++, date.c_str(), NULL);
        worksheet_write_string(worksheet, row, col++, time.c_str(), NULL);

        // Write settings
        worksheet_write_string(worksheet, row, col++, safe_get_string(settings, "changeSuit").c_str(), NULL);
        worksheet_write_number(worksheet, row, col++, safe_get_int(settings, "gameOfScoreChange"), NULL);
        worksheet_write_boolean(worksheet, row, col++, safe_get_bool(settings, "includeTimer"), NULL);
        worksheet_write_number(worksheet, row, col++, safe_get_int(settings, "maxTimePerRound"), NULL);
        worksheet_write_number(worksheet, row, col++, num_games, NULL);
        worksheet_write_number(worksheet, row, col++, safe_get_int(settings, "roundOfScoreChange"), NULL);

        // Loop through games
        for (int game = 1; game <= num_games; ++game) {
            auto game_moves = participant_data["moves"]["Game_" + std::to_string(game)]["moves"];

            auto finalKey = participant_data["stats"]["Game_" + std::to_string(game)]["FinalScore"];

            int moveIndex = 0;
            // Loop through rounds (1 to 10) for each game
            for (int round = 1; round <= 10; ++round) {
               
                // Loop through moves for each round (up to 10 moves per round)
                for (int move = 1; move <= 8; ++move) {
                    auto individual_move = participant_data["moves"]["Game_" + std::to_string(game)]["moves"][moveIndex];
                    // Ensure the move index is within the range of the current round's max moves
                    worksheet_write_number(worksheet, row, col++, safe_get_int(individual_move, "moveNumber"), NULL);
                    worksheet_write_string(worksheet, row, col++, safe_get_string(individual_move, "points").c_str(), NULL);
                    worksheet_write_number(worksheet, row, col++, safe_get_int(individual_move, "time"), NULL);
                    worksheet_write_string(worksheet, row, col++, safe_get_string(individual_move, "type").c_str(), NULL);
                    worksheet_write_string(worksheet, row, col++, safe_get_string(individual_move, "row").c_str(), NULL);
                    worksheet_write_string(worksheet, row, col++, safe_get_string(individual_move, "col").c_str(), NULL);
                    worksheet_write_string(worksheet, row, col++, safe_get_string(individual_move, "rank").c_str(), NULL);
                    worksheet_write_string(worksheet, row, col++, safe_get_string(individual_move, "suit").c_str(), NULL);

                    moveIndex++;
                }

                // Write the stats for this round
                auto totalScoreKey = participant_data["stats"]["Game_" + std::to_string(game)]["TotalScore_Round_" + std::to_string(round)];
                auto exploitationKey = participant_data["stats"]["Game_" + std::to_string(game)]["Exploitation_Round_" + std::to_string(round)];
                auto explorationKey = participant_data["stats"]["Game_" + std::to_string(game)]["Exploration_Round_" + std::to_string(round)];
                auto unexploredKey = participant_data["stats"]["Game_" + std::to_string(game)]["PercentUnexplored_Round_" + std::to_string(round)];
                auto avgScoreKey = participant_data["stats"]["Game_" + std::to_string(game)]["AvgTotalScore_Round_" + std::to_string(round)];
                worksheet_write_number(worksheet, row, col++, safe_get_double(totalScoreKey, "totalScore"), NULL);
                worksheet_write_number(worksheet, row, col++, safe_get_int(exploitationKey, "exploitativeMoves"), NULL);
                worksheet_write_number(worksheet, row, col++, safe_get_int(explorationKey, "exploratoryMoves"), NULL);
                worksheet_write_number(worksheet, row, col++, safe_get_double(unexploredKey, "percentUnexplored"), NULL);
                worksheet_write_number(worksheet, row, col++, safe_get_double(avgScoreKey, "averageTotalScores"), NULL);
            }

            worksheet_write_number(worksheet, row, col++, safe_get_int(finalKey, "TotalScoreAcrossGames"), NULL);

        }

        // Move to the next row for the next participant
        row++;
    }


    /*
    // Columns for moves and stats
    int move_col_base = 9;
    int stats_col_base = 9 + 10 * 6;  // Each round has 10 moves (6 columns per move)

    // Add columns for 10 rounds of moves
    for (int round = 1; round <= 10; ++round) {
        for (int move = 1; move <= 10; ++move) {
            worksheet_write_string(worksheet, 0, move_col_base++,
                ("Move_" + std::to_string(round) + "_" + std::to_string(move) + "_Number").c_str(), NULL);
            worksheet_write_string(worksheet, 0, move_col_base++,
                ("Move_" + std::to_string(round) + "_" + std::to_string(move) + "_Points").c_str(), NULL);
            worksheet_write_string(worksheet, 0, move_col_base++,
                ("Move_" + std::to_string(round) + "_" + std::to_string(move) + "_Rank").c_str(), NULL);
            worksheet_write_string(worksheet, 0, move_col_base++,
                ("Move_" + std::to_string(round) + "_" + std::to_string(move) + "_Suit").c_str(), NULL);
            worksheet_write_string(worksheet, 0, move_col_base++,
                ("Move_" + std::to_string(round) + "_" + std::to_string(move) + "_Time").c_str(), NULL);
            worksheet_write_string(worksheet, 0, move_col_base++,
                ("Move_" + std::to_string(round) + "_" + std::to_string(move) + "_Type").c_str(), NULL);
        }
    }

    // Add columns for stats for 10 rounds
    for (int round = 1; round <= 10; ++round) {
        worksheet_write_string(worksheet, 0, stats_col_base++,
            ("AvgTotalScore_Round_" + std::to_string(round)).c_str(), NULL);
        worksheet_write_string(worksheet, 0, stats_col_base++,
            ("Exploitation_Round_" + std::to_string(round)).c_str(), NULL);
        worksheet_write_string(worksheet, 0, stats_col_base++,
            ("Exploration_Round_" + std::to_string(round)).c_str(), NULL);
        worksheet_write_string(worksheet, 0, stats_col_base++,
            ("PercentUnexplored_Round_" + std::to_string(round)).c_str(), NULL);
    }

    // Add column for total score
    worksheet_write_string(worksheet, 0, stats_col_base++, "TotalScore", NULL);

    int row = 1;

    // Loop through participants
    for (auto& participant : data["participants"].items()) {
        auto participant_data = participant.value();
        std::string participant_id = participant.key();
        std::string date = safe_get_string(participant_data["datetime"], "date");
        std::string time = safe_get_string(participant_data["datetime"], "time");
        auto settings = participant_data["settings"];
        int num_games = safe_get_int(settings, "numGames");

        // Write participant info (ID, Date, Time)
        worksheet_write_string(worksheet, row, 0, participant_id.c_str(), NULL);
        worksheet_write_string(worksheet, row, 1, date.c_str(), NULL);
        worksheet_write_string(worksheet, row, 2, time.c_str(), NULL);

        // Write settings
        worksheet_write_string(worksheet, row, 3, safe_get_string(settings, "changeSuit").c_str(), NULL);
        worksheet_write_number(worksheet, row, 4, safe_get_int(settings, "gameOfScoreChange"), NULL);
        worksheet_write_boolean(worksheet, row, 5, safe_get_bool(settings, "includeTimer"), NULL);
        worksheet_write_number(worksheet, row, 6, safe_get_int(settings, "maxTimePerRound"), NULL);
        worksheet_write_number(worksheet, row, 7, num_games, NULL);
        worksheet_write_number(worksheet, row, 8, safe_get_int(settings, "roundOfScoreChange"), NULL);

        int move_col = move_col_base;
        int stats_col = stats_col_base;
        double total_score = 0; // Variable for total score

        // Loop through games
        for (int game = 1; game <= num_games; ++game) {
            auto game_moves = participant_data["moves"]["Game_" + std::to_string(game)]["moves"];
            double game_score = 0; // Variable for each game's score

            // Loop through rounds (1 to 10) for each game
            for (int round = 1; round <= 10; ++round) {
                int move_index = 1;

                // Loop through moves for each round (up to 10 moves per round)
                for (auto& move : game_moves) {
                    // Ensure the move index is within the range of the current round's max moves
                    if (move_index > 10) break;
                    worksheet_write_number(worksheet, row, move_col++, safe_get_int(move, "moveNumber"), NULL);
                    worksheet_write_string(worksheet, row, move_col++, safe_get_string(move, "points").c_str(), NULL);
                    worksheet_write_string(worksheet, row, move_col++, safe_get_string(move, "rank").c_str(), NULL);
                    worksheet_write_string(worksheet, row, move_col++, safe_get_string(move, "suit").c_str(), NULL);
                    worksheet_write_number(worksheet, row, move_col++, safe_get_int(move, "time"), NULL);
                    worksheet_write_string(worksheet, row, move_col++, safe_get_string(move, "type").c_str(), NULL);
                    ++move_index;

                    // Accumulate game score based on points or any other scoring logic
                    game_score += std::stoi(move["points"].get<std::string>());
                }

                // Write the stats for this round
                std::string stats_key = "Game_" + std::to_string(game);
                worksheet_write_number(worksheet, row, stats_col++, safe_get_double(participant_data["stats"][stats_key], "AvgTotalScore_Round_" + std::to_string(round)), NULL);
                worksheet_write_number(worksheet, row, stats_col++, safe_get_int(participant_data["stats"][stats_key], "Exploitation_Round_" + std::to_string(round)), NULL);
                worksheet_write_number(worksheet, row, stats_col++, safe_get_int(participant_data["stats"][stats_key], "Exploration_Round_" + std::to_string(round)), NULL);
                worksheet_write_number(worksheet, row, stats_col++, safe_get_double(participant_data["stats"][stats_key], "PercentUnexplored_Round_" + std::to_string(round)), NULL);
            }

            // Add total score for the game
            total_score += game_score; // Accumulate score for all games
        }

        // Write the total score for the participant
        worksheet_write_number(worksheet, row, stats_col++, total_score, NULL);

        // Move to the next row for the next participant
        row++;
    }

    */

    workbook_close(workbook);
    MessageBox(NULL, L"Conversion complete.", L"Success", MB_OK | MB_ICONEXCLAMATION);
}







void beginConvert(std::wstring filePath) {
    if (!filePath.empty()) {
        // File was selected, now load the JSON file
        std::ifstream inputFile(filePath);
        if (inputFile.is_open()) {
            nlohmann::json data;
            inputFile >> data;  // Load JSON data from file
            createExcel(data);  // Convert the JSON data to Excel
        }
        else {
            MessageBox(NULL, L"Failed to open the file.", L"Error", MB_OK | MB_ICONERROR);
        }
    }
}

#define MAX_LOADSTRING 100

// Global Variables:
HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name

// Forward declarations of functions included in this code module:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // TODO: Place code here.

    // Initialize global strings
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_DEEPERPATHSCONVERTER, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // Perform application initialization:
    if (!InitInstance (hInstance, nCmdShow))
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_DEEPERPATHSCONVERTER));

    MSG msg;

    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return (int) msg.wParam;
}



//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_DEEPERPATHSCONVERTER));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_DEEPERPATHSCONVERTER);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance; // Store instance handle in our global variable

   HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

   if (!hWnd)
   {
      return FALSE;
   }

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE: Processes messages for the main window.
//
//  WM_COMMAND  - process the application menu
//  WM_PAINT    - Paint the main window
//  WM_DESTROY  - post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_COMMAND:
        {
            int wmId = LOWORD(wParam);
            // Parse the menu selections:
            switch (wmId)
            {
            case IDM_ABOUT:
                DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
                break;
            case IDM_CONVERT: // Menu option for file selection and conversion
            {
                std::wstring selectedFile = OpenFileDialog();
                if (!selectedFile.empty()) {
                    MessageBox(hWnd, selectedFile.c_str(), L"Selected File", MB_OK);
                    beginConvert(selectedFile);
                }
                else {
                    MessageBox(hWnd, L"No file selected.", L"Error", MB_OK);
                }
            }
                break;
            case IDM_EXIT:
                DestroyWindow(hWnd);
                break;
            default:
                return DefWindowProc(hWnd, message, wParam, lParam);
            }
        }
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            // TODO: Add any drawing code that uses hdc here...
            EndPaint(hWnd, &ps);
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}
