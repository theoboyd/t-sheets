#(c) 2010-2011 Theodore Boyd
import sys

# Globals
cells = [[]]  # 3D array of chars / 2D array of strings
editing_mode = False

sel_options = {'x': 0, 'y': 0, 'col': 0, 'row': 0, 'all': False}
formula_sel_options = {'x': 0, 'y': 0, 'start_x': 0, 'start_y': 0, 'end_x': 0, 'end_y': 0}
grid_options = {'width': 80, 'height': 23, 'cell_width': 7, 'cell_height': 3}

file = './t-sheets.dat'  # Default file to save to (in program dir)
copyright_and_status = 'Ready. Type \'h\' for help. | Theodore Boyd 2011 | T-Sheets Version 1.4'
end_of_grid = 'Reached the end of the grid.'
formula = 0
formulae = ['SUM', 'AVG']

def main(*args):
  global grid_options

  if sys.version_info[0] < 3 or (sys.version_info[0] == 3 and sys.version_info[1] < 1):
    print('You are using an unsupported version of Python.\nTry running with the command \'python3.1 t-sheets.py\'.')
    return

  launch_shortcut = 'r'  # The shortuct for starting T-Sheets
  
  preparecells((grid_options['width'] // grid_options['cell_width']) - 1, (grid_options['height'] - prepuitop() - prepuibottom()))
  
  response = ''
  unexpected_flag = ''
  try:
    flag = args[1]
  except:
    flag = ''
  
  if flag == '-n' or flag == '--no-prompt':
    # Don't display resize info / warning
    response = launch_shortcut
  else:
    if flag:
      unexpected_flag = '\n \n \n \n(Unexpected flag %s was ignored.)' % flag
    messagebox('Please resize your window so that this box fills it exactly.\nLeave one extra line at the bottom for the prompt.\n \nType \'r\' when ready, or \'q\' to quit.%s' % unexpected_flag, ['Quit (q)', 'Ready (' + launch_shortcut + ')'])
    response = getvalidchar(['q', launch_shortcut])
  
  if response == launch_shortcut:
    launchui()
  elif response == 'q':
    return

def launchui(footermessage=copyright_and_status, forcemenu=None):
  global cells, grid_options, editing_mode, end_of_grid, sel_options
  drawuitop()
  drawgrid(grid_options['width'], grid_options['height'] - prepuitop() - prepuibottom())
  drawuibottom(footermessage)
  
  if editing_mode and not forcemenu:
    string = getandshowstring(['`'])
    if len(string) > 0:
      setcell(sel_options['x'], sel_options['y'], string.strip())  # Save space, and allow for clearing cells.
    editing_mode = False
    launchui(copyright_and_status)
  else:
    if forcemenu:
      response = forcemenu
    else:
      response = getvalidchar(['f', 'e', 'v', 'h', 'q', '`', 'w', 'a', 's', 'd', 'r', 'c', 'l', 'x'])
    if response == 'f':
      global file
      drawuitop([['File', 'f', True], ['Edit', 'e', False], ['View', 'v', False], ['Help', 'h', False]])
      messagebox('Open or save spreadsheet file:\n \n \n \nFile: [ ' + file + ' ]\n \nWarning: if this file exists, saving will overwrite it.', ['Open (o)', 'Save To (s)', 'Specify File (`)'], True)
      fileaction = getvalidchar(['o', 's', '`', 'f'])  # Wait for shortcut to be pressed
      if fileaction == 'o':
        result = loadgridfromfile(file)
        launchui(result)
      elif fileaction == 's':
        result = savegridtofile(file)
        launchui(result)
      elif fileaction == '`':
        file = getandshowstring(['`'], 'File: [ ', ' ]', True)
        launchui(copyright_and_status, 'f')
      else:
        launchui()
    elif response == 'e':
      drawuitop([['File', 'f', False], ['Edit', 'e', True], ['View', 'v', False], ['Help', 'h', False]])
      messagebox('To edit, use \'`\' (backquote).', ['Insert Forumla (i)'], True)
      fileaction = getvalidchar(['i', 'e'])
      if fileaction == 'i':
        formulaeditor()
      else:
        launchui()
    elif response == 'v':
      drawuitop([['File', 'f', False], ['Edit', 'e', False], ['View', 'v', True], ['Help', 'h', False]])
      messagebox('No view options available.', [], True)
      getvalidchar(['v'])
      launchui()
    elif response == 'h':
      drawuitop([['File', 'f', False], ['Edit', 'e', False], ['View', 'v', False], ['Help', 'h', True]])
      messagebox('Run (UNIX):\n\'python t-sheets.py [-n]\', where \'python\' is the name of the Python 3+ executable.\n-n, --no-prompt: Skip the prompt to resize the window.\n \nNavigation:\nUse the \'w\', \'a\', \'s\' and \'d\' keys on your keyboard to move through the grid and \'`\' (backquote) to toggle editing a cell. Or select (r)ows, (c)olumns, a(l)l or specific cells (x).\n \nFormulae:\nBegin a cell with an \'=\' sign to perform addition, subtraction, multiplication or division. For more advanced formulae, use the formula editor (i) under the Edit (e) menu.', [], True)
      getvalidchar(['h'])
      launchui()
    elif response == 'q':
      return
    elif response == '`':
      editing_mode = True
      xy = str(cells[sel_options['x']][sel_options['y']])
      launchui('Editing mode. | Cell = \'' + xy + '\'. Type text to replace, space to clear.')
    elif response == 'w':
      if sel_options['y'] > 0:
        sel_options['y'] -= 1
        clearsel_exceptcells()
        launchui(copyright_and_status)
      else:
        launchui(end_of_grid)
    elif response == 'a':
      if sel_options['x'] > 0:
        sel_options['x'] -= 1
        clearsel_exceptcells()
        launchui(copyright_and_status)
      else:
        launchui(end_of_grid)
    elif response == 's':
      if sel_options['y'] < (grid_options['height'] - prepuitop() - prepuibottom() - 2):
        sel_options['y'] += 1
        clearsel_exceptcells()
        launchui(copyright_and_status)
      else:
        launchui(end_of_grid)
    elif response == 'd':
      if sel_options['x'] < ((grid_options['width'] // grid_options['cell_width']) - 2):
        sel_options['x'] += 1
        clearsel_exceptcells()
        launchui(copyright_and_status)
      else:
        launchui(end_of_grid)
    elif response == 'r':
      rownum = getandshowstring(['`'], 'Select row #', '', True)
      try:
        selectrow(int(rownum))
      except:
        pass
      launchui(copyright_and_status)
    elif response == 'c':
      colnum = getandshowstring(['`'], 'Select column #', '', True)
      try:
        selectcol(int(colnum))
      except:
        pass
      launchui(copyright_and_status)
    elif response == 'l':
      selectall()
      launchui(copyright_and_status)
    elif response == 'x':
      colnum = getandshowstring([','], 'Select cell (', ', _)', True)
      rownum = getandshowstring(['`'], 'Select cell (' + colnum + ', ', ')', True)
      try:
        selectcell(int(colnum), int(rownum))
      except:
        pass
      launchui(copyright_and_status)

def formulaeditor():
  global formula_sel_options, formulae, formula
  
  #formulatask('Select the first cell in the range using 'w', 'a', 's' and 'd'.\n \nThen type '`' to confirm.', stage)
  
  # Select first cell in range using WASD
  chooseformulacell('Select the first cell in the range.')
  formula_sel_options['start_x'] = formula_sel_options['x']
  formula_sel_options['start_y'] = formula_sel_options['y']

  # Use WASD to enlarge selection range
  chooseformulacell('Now enlarge the selection range if necessary.', formula_sel_options['start_x'], formula_sel_options['start_y'])
  formula_sel_options['end_x'] = formula_sel_options['x']
  formula_sel_options['end_y'] = formula_sel_options['y']

  # Select formula to apply to range using WS
  chooseformula('Select the formula to apply.')

  # Calculate and display answer
  val = 0
  if formula_sel_options['start_x'] > formula_sel_options['end_x']:
    temp = formula_sel_options['start_x']
    formula_sel_options['start_x'] = formula_sel_options['end_x']
    formula_sel_options['end_x'] = temp
  if formula_sel_options['start_y'] > formula_sel_options['end_y']:
    temp = formula_sel_options['start_y']
    formula_sel_options['start_y'] = formula_sel_options['end_y']
    formula_sel_options['end_y'] = temp

  if formula == 0:
    # Sum
    for x in range(formula_sel_options['start_x'], formula_sel_options['end_x']):
      for y in range(formula_sel_options['start_y'], formula_sel_options['end_y']):
        try:
          val += float(cells[x][y])
          messagebox(str(val), [])
        except:
          val = 'Error #3A: Could not add value of cell (%s, %s) to sum' % (x + 1, y + 1)
  elif formula == 1:
    # Average
    sum = 0
    tot = 0
    for x in range(formula_sel_options['start_x'], formula_sel_options['end_x']):
      for y in range(formula_sel_options['start_y'], formula_sel_options['end_y']):
        try:
          sum += float(cells[x][y])
          tot += 1
          messagebox(str(tot), [])
        except:
          val = 'Error #3B: Could not add value of cell (%s, %s) to sum' % (x + 1, y + 1)
    try:
      val = sum / tot
    except:
      val = 'Error #3C: Could not divide sum (%s) by total (%s)' % (sum, tot)
  else:
    val = 'Error #3D: Unexpected formula'
  
  # Optional text form of formula
  #valtext = '=' + formulae[formula] + '(' + str(formula_startx) + ',' + str(formula_starty) + ':' + str(formula_endx) + ',' + str(formula_endy) + ')'
  
  # Select cell to place answer using WASD
  # Defaults to first cell in range
  chooseformulacell('Select the cell in which to save the answer.')
  setcell(formula_sel_options['x'], formula_sel_options['y'], str(val))
  launchui()
  
def chooseformulacell(message, anchorx = -1, anchory = -1):
  global grid_options, formula_sel_options
  response = ''
  oldscalex = ((grid_options['width'] // grid_options['cell_width']) - 1)
  oldscaley = (grid_options['height'] - prepuitop() - prepuibottom() - 1)
  while response != '`':
    minigrid = '\n  '
    for x in range(0, oldscalex):
      minigrid += '  '
    minigrid += '   Formula Editor\n '
    for y in range(0, oldscaley):
      minigrid += ' : '
      for x in range(0, oldscalex):
        if formula_sel_options['x'] == x and formula_sel_options['y'] == y:
          minigrid += '+ '
        else:
          if anchorx != -1 and anchory != -1:
            if anchorx == x and anchory == y:
              minigrid += '+ '
            elif (
                (anchorx <= x and x <= formula_sel_options['x'])
                and
                (anchory <= y and y <= formula_sel_options['y'])
               ) or (
                (anchorx <= x and x <= formula_sel_options['x'])
                and
                (formula_sel_options['y'] <= y and y <= anchory)
               ) or (
                (formula_sel_options['x'] <= x and x <= anchorx)
                and
                (anchory <= y and y <= formula_sel_options['y'])
               ) or (
                (formula_sel_options['x'] <= x and x <= anchorx)
                and
                (formula_sel_options['y'] <= y and y <= anchory)
               ):
              # Cells are within the bounding box formed by
              # the original point and the new point
              minigrid += '. '
            else:
              try:
                minigrid += cells[x][y][0] + ' '
              except:
                minigrid += '  '
          else:
            try:
              minigrid += cells[x][y][0] + ' '
            except:
              minigrid += '  '
      minigrid += ': '
      if y == 3:
        minigrid += ' ' + message
      if y == 5:
        minigrid += ' Use the \'w\', \'a\', \'s\' and \'d\' keys to navigate.'
      minigrid += '\n '
  
    drawuitop([['File', 'f', False], ['Edit', 'e', True], ['View', 'v', False], ['Help', 'h', False]])
    messagebox(minigrid, ['Confirm (' + str(formula_sel_options['x'] + 1) + ', ' + str(formula_sel_options['y'] + 1) +  ') (`)'])
    response = getvalidchar(['`', 'w', 'a', 's', 'd'])
    if response == 'w':
      if formula_sel_options['y'] > 0:
        formula_sel_options['y'] -= 1
    elif response == 'a':
      if formula_sel_options['x'] > 0:
        formula_sel_options['x'] -= 1
    elif response == 's':
      if formula_sel_options['y'] < (grid_options['height'] - prepuitop() - prepuibottom() - 2):
        formula_sel_options['y'] += 1
    elif response == 'd':
      if formula_sel_options['x'] < ((grid_options['width'] // grid_options['cell_width']) - 2):
        formula_sel_options['x'] += 1
  
def chooseformula(message):
  global formula, formulae
  response = ''
  while response != '`':
    formulalist = ''
    for i in range (0, len(formulae)):
      front = ' '
      back = ' '
      if i == formula:
        front = '['
        back = ']'
      formulalist += front + formulae[i] + back + '\n'
    drawuitop([['File', 'f', False], ['Edit', 'e', True], ['View', 'v', False], ['Help', 'h', False]])
    messagebox('Formula Editor\n \n' + message + '\n \n' + formulalist + '\n \nChoose with \'w\' and \'s\'.', ['Confirm ' + formulae[formula] +' (`)'])
    response = getvalidchar(['`', 'w', 's'])
    if response == 'w':
      if formula > 0:
        formula -= 1
    elif response == 's':
      if formula < len(formulae) - 1:
        formula += 1

def formulatask(taskmessage, currentstage):
  drawuitop([['File', 'f', False], ['Edit', 'e', True], ['View', 'v', False], ['Help', 'h', False]])
  messagebox('Formula Editor\n \n' + taskmessage, ['Cancel (c)', 'OK (o)'], True)
  response = getvalidchar(['o', 'c', 'e'])
  if response == 'o':
    formulaeditor(currentstage + 1)
  elif response == 'c':
    launchui(copyright_and_status, 'e')
  elif response == 'e':
    launchui()

def loadgridfromfile(path):
  try:
    f = open(path, 'r')
    rows = f.readlines()
    for i in range(0, len(rows)):
      cells = rows[i].split('`')
      for j in range (0, len(cells) - 1):
        setcell(i, j, cells[j])
    f.close
    return 'Opened ' + path
  except:
    return 'File could not be opened or does not exist.'

def savegridtofile(path):
  global grid_options
  try:
    f = open(path, 'w')
    for i in range(0, (grid_options['width'] // grid_options['cell_width']) - 1):
      for j in range(0, grid_options['height'] - prepuitop() - prepuibottom()):
        f.write(cells[i][j] + '`')
      f.write('\n')
    f.close
    return 'Saved to ' + path
  except IOError as err:
    return 'File could not be saved (' + str(err)[10:grid_options['width'] - 24] + '_).'

def clearsel_exceptcells():
  global sel_options
  sel_options['col'] = 0
  sel_options['row'] = 0
  sel_options['all'] = False

def selectcell(x, y):
  global grid_options, sel_options
  if (x > 0 and x <= ((grid_options['width'] // grid_options['cell_width']) - 1)) and (y > 0 and y <= (grid_options['height'] - prepuitop() - prepuibottom() - 1)):
    sel_options['x'] = x - 1
    sel_options['y'] = y - 1
  # Else, selected cell is unchanged
  sel_options['col'] = 0
  sel_options['row'] = 0
  sel_options['all'] = False
  
def selectall():
  global sel_options
  sel_options['x'] = 0
  sel_options['y'] = 0
  sel_options['col'] = 0
  sel_options['row'] = 0
  sel_options['all'] = not sel_options['all']  # This toggles
  
def selectcol(c = 0):
  global sel_options
  sel_options['x'] = 0
  sel_options['y'] = 0
  sel_options['col'] = c
  sel_options['row'] = 0
  sel_options['all'] = False
  
def selectrow(r = 0):
  global sel_options
  sel_options['x'] = 0
  sel_options['y'] = 0
  sel_options['col'] = 0
  sel_options['row'] = r
  sel_options['all'] = False

def preparecells(x, y, val = ''):
  global cells
  for i in range(0, x):
    cells.insert(i, [val])
  for i in range(0, x):
    for j in range(0, y):
      cells[i].insert(j, val)
    
def setcell(x, y, val):
  global cells
  try:
    cells[x][y] = val
  except:
    pass
    #cells.insert(x, cells[x].insert(y, val))
    
def drawgrid(lenx, leny):
  global cells, grid_options, sel_options
  totx = lenx // grid_options['cell_width']
  toty = (leny - 0) # // grid_options['cell_height'] #Padding of 4
  selAll = '(rclx)'
  for y in range(0, toty):
    print('| ', end = '')
    for x in range(0, totx):
      try:
        xy = cells[x - 1][y - 1]
      except:
        xy = ''
      if y == 0 and x == 0:
        print(selAll + (' ' * (grid_options['cell_width'] - 1 - len(selAll)) + ' '), end = '')
      elif x == 0:
        print(str(y) + (':' * (grid_options['cell_width'] - 1 - len(str(y))) + ' '), end = '')
      elif y == 0:
        print(str(x) + (':' * (grid_options['cell_width'] - 1 - len(str(x))) + ' '), end = '')
      else:
        left = ''
        right = ''
        sel = 0
        if sel_options['all'] or (y == sel_options['y'] + 1 and x == sel_options['x'] + 1) or (sel_options['row'] > 0 and y == sel_options['row']) or (sel_options['col'] > 0 and x == sel_options['col']):
          left = '['
          right = ']'
          sel = 2
        
        elipsis = ''
        xy = parseequation(xy)
        if len(xy) > (grid_options['cell_width'] - 1 - sel):
          elipsis = '_'
          #display = xy[0:grid_options['cell_width'] - 3 - sel] + '..'
        #else:
        display = str(xy[0:grid_options['cell_width'] - 1 - len(elipsis) - sel]) + elipsis
        print(left + display + (' ' * (grid_options['cell_width'] - 1 - sel - len(display))) + right + ' ', end = '')
    print('|')

def parseequation(string):
  try:
    if string[0] == '=' and string[1] != '=':  # Implicitly checks string[1] exists
      try:
        (lhs, rhs, token) = consumetotoken(string[1:], ['+','-','*','/'])
        if token == '+':
          return str(float(lhs) + float(rhs))
        elif token == '-':
          return str(float(lhs) - float(rhs))
        elif token == '*':
          return str(float(lhs) * float(rhs))
        elif token == '/':
          return str(float(lhs) / float(rhs))
        else:  # End of line
          return lhs
      except:
        return 'Error #30: Could not parse arithmetic equation'
    else:
      return string
  except:
    return string

def consumetotoken(string, tokenlist):
  lhs = string
  rhs = ''
  token = ''
  for i in range(0, len(string) - 1):
    if string[i] in tokenlist:
      token = string[i]
      lhs = string[0:i]
      rhs = string[(i + 1):]
      i = len(string)  # Break out

  return (lhs, rhs, token)

def editingmode(escapechars = ['\n']):
  # Enter this mode to edit text and not activate f, e, v, h or q
  pass

def launchhelp(query = ''):
  global grid_options
  newheight = grid_options['height'] - prepuitop()
  drawuitop()
  for i in range(0, newheight - prepuibottom()): drawline(grid_options['width'], [], ' ', ' ', '|')
  if query == '':
    drawuibottom('Please type your query followed by a \'?\'.')
  else:
    drawuibottom('\'' + query + '\'')
  return getstring('?')

def getvalidchar(validchars = []):
  response = ''
  while response not in validchars:
    response = getchar()
  return response

def getandshowstring(endofstring = ['`'], displeft = '\'', dispright = '\'', instanton = False):
  global grid_options
  string = ''
  if instanton:
    drawuitop()
    drawgrid(grid_options['width'], grid_options['height'] - prepuitop() - prepuibottom())
    drawuibottom(displeft + string + dispright + '. Type ' + str(endofstring) + ' to finish input.')
  c = ''
  c = getchar()
  while c not in endofstring:
    string += c
    drawuitop()
    drawgrid(grid_options['width'], grid_options['height'] - prepuitop() - prepuibottom())
    drawuibottom(displeft + string + dispright + '. Type ' + str(endofstring) + ' to finish input.')
    c = getchar()
  return string

def getstring(endofstring = ['`']):
  string = ''
  c = ''
  c = getchar()
  while c not in endofstring:
    string += c
    c = getchar()
  return string  

def drawline(markers = [], style = '-', markerstyle = '+', sidestyle = '+', intramarker = '-', useintramarkers = False):
  global grid_options
  intramarkermode = False
  print(sidestyle, end = '')
  if markers == []:
    print(style * (grid_options['width'] - 2), end = '')
  else:
    for i in range(0, grid_options['width'] - 2):
      if (i - 1) in markers:
        print(markerstyle, end = '')
        intramarkermode = not intramarkermode
      else:
        if useintramarkers and intramarkermode:
          print(intramarker, end = '')
        else:
          print(style, end = '')
  print(sidestyle)
  return 1

def drawstring(string, leftspace = ' '):
  global grid_options
  rightspace = ''
  printedlines = 0
  totstring = len(string)
  remaining = totstring
  if len(leftspace) != 1:
    rightspace = ' '
  while remaining > 0:
    if remaining < (grid_options['width'] - 4):
      print('|' + leftspace + string[(totstring - remaining):(totstring - remaining + grid_options['width'] - 4)] + rightspace, end = '')
      print((' ' * (grid_options['width'] - 4 - remaining)) + ' |')
    else:
      print('|' + leftspace + string[(totstring - remaining):(totstring - remaining + grid_options['width'] - 4)] + rightspace + ' |')
    remaining -= (grid_options['width'] - 4)
    printedlines += 1
  return printedlines
     
def messagebox(message, buttons, customtop = False):
  global grid_options
  messages = message.split('\n')
  newheight = grid_options['height']
  if customtop:
    newheight -= prepuitop()
  else:
    newheight -= drawline()
  for string in messages: newheight -= drawstring(string)
  for i in range(0, newheight - prepbuttonstrip() - 1): drawline([], ' ', ' ', '|')
  drawbuttonstrip(*buttons)
  drawline()
  
def drawuitop(items=[['File', 'f', False], ['Edit', 'e', False], ['View', 'v', False], ['Help', 'h', False]]):
  displaystring = ''
  blockwidth = 10

  for item in items:
    block = ''
    selected = item[2]
    anotherselected = False
    controlhint = ' (' + item[1] + ')'
    if selected:
      block += '[' + item[0] + controlhint
    else:
      block += ' ' + item[0]
      for item in items:
        if item[2]:
          anotherselected = True
          break
      if not anotherselected:
        block += controlhint

    if selected:
      block +=  (' ' * (blockwidth - len(block) - 1)) + ']|'
    else:
      block +=  (' ' * (blockwidth - len(block))) + '|'
    displaystring += block

  drawline([9, 20, 31, 42])
  drawstring(displaystring, '')
  drawline([9, 20, 31, 42])

def prepuitop():
  return 3

def drawuibottom(message):
  drawline()
  drawstring(message)
  drawline()

def prepuibottom():
  return 3  # Needed to preempt how tall the bottom of the UI will be

def drawbuttonstrip(*buttons):
  global grid_options
  topstring = ' '
  middlestring = ' '
  bottomstring = ' '
  for buttontext in buttons:
    topstring +=    '+-' + ('-' * len(buttontext)) + '-+  '
    middlestring += '| ' +        buttontext       + ' |  '
    bottomstring += '+-' + ('-' * len(buttontext)) + '-+  '
  remaining = grid_options['width'] - len(middlestring) - 3
  print('| ' + (' ' * remaining) + topstring + '|')
  print('| ' + (' ' * remaining) + middlestring + '|')
  print('| ' + (' ' * remaining) + bottomstring + '|')
  
def prepbuttonstrip():
  return 3  # Needed to preempt how tall the button will be

def getchar():
   # Returns a single character from standard input
   import tty, termios
   fd = sys.stdin.fileno()
   old_settings = termios.tcgetattr(fd)
   try:
    tty.setraw(sys.stdin.fileno())
    ch = sys.stdin.read(1)
   finally:
    termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)
   return ch

if __name__ == '__main__':
  sys.exit(main(*sys.argv))
