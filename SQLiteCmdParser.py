import sys
import os.path
import sqlite3

def SQLiteCmdParser(Db_path, CmdFile_path):
    if os.path.exists(Db_path) == False:
        return
        
    if os.path.exists(CmdFile_path) == False:
        return
        
        
    conn = sqlite3.connect(Db_path)
    c = conn.cursor()

    cmd_file = open(CmdFile_path, 'r')
    cmd = ""
    
    for line in cmd_file:
        line = line.strip("\n").strip(" ")
        cmd = cmd + " " + line
        
        if cmd[-1] == ";":
            c.execute(cmd)
            cmd = ""
        

    conn.commit()
    conn.close()


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit()

    SQLiteCmdParser(sys.argv[1], sys.argv[2])
    


