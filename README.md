
# Tic Tac Toe in Excel using VBA

This is a simple implementation of the classic Tic Tac Toe game built using Microsoft Excel and VBA (Visual Basic for Applications). The game allows two players to play on a 3x3 grid directly in Excel by double-clicking cells to mark Xs and Os. This project showcases how VBA macros can be used to create interactive applications within Excel.

## Features

- **Interactive Gameplay**: Play directly in Excel by double-clicking cells.
- **Automatic Turn Switching**: Automatically switches turns between Player X and Player O.
- **Win Detection**: Detects win conditions (rows, columns, diagonals) and announces the winner.
- **Draw Detection**: Recognizes when the game ends in a draw.
- **Restart Functionality**: Automatically restarts the game after a win or draw.

## How to Set Up

1. **Download the Project**: Clone or download the repository from GitHub.

   ```bash
   git clone https://github.com/prakashtaz0091/Tic-Tac-Toe.git
2. **Open Excel:** Open the Excel file (TicTacToe.xltm) in Microsoft Excel.

3. **Enable Macros:** Make sure to enable macros in Excel when prompted, as the game logic relies on VBA scripts.

## How to Play
1. **Initialize the Game:**

Press the Start Game button at the top or Alt + F8, select InitializeGame, and click Run to start a new game. This will clear the board and set the game to start with Player X.
2. **Play the Game:**

Double-click any cell within the 3x3 grid (D4
) to mark your move.
Player X always starts, and players take turns marking cells with X or O.

3. **Winning the Game:**

The game checks for three consecutive marks (X or O) in a row, column, or diagonal. If a player wins, a message box will announce the winner, and the game will automatically restart.

4. **Draw:**

If all cells are filled without a winner, the game declares a draw and restarts.


## File Structure
1. **VBA Code Location:**
All code is placed inside Sheet1, which is renamed as "Tic Tac Toe". This includes the InitializeGame subroutine and the Worksheet_BeforeDoubleClick event handling the game logic.

## Troubleshooting
 1. Ensure that the Worksheet_BeforeDoubleClick event code is correctly placed inside the "Tic Tac Toe" sheet object in the VBA editor.
2. Make sure macros are enabled; otherwise, the game will not function correctly.
3. If the cells do not respond, check that the correct range (D4
) is being double-clicked.

## Contributing
If you have suggestions for improvements or want to contribute, feel free to fork the repository and submit pull requests. All contributions are welcome!

## License
This project is open source and available under the MIT License.
