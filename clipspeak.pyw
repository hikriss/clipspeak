import wx
import win32api
import win32gui
import win32con
import win32clipboard
import win32com.client

Engine = win32com.client.Dispatch('SAPI.SPVoice')

def create(parent):
    return ClipMonFrame(parent)
# assign ID numbers
[wxID_FRAME1, wxID_FRAME1BTN_CLEAR, wxID_FRAME1BTN_CLEARALL, wxID_FRAME1LB_CBLIST,] = [wx.NewId() for _init_ctrls in range(4)]
   
class ClipMonFrame (wx.Frame):
    def _init_ctrls(self, prnt):
        # BOA generated methods
        wx.Frame.__init__(self, id=wxID_FRAME1, name='', parent=prnt, size=wx.Size(300, 400), style=wx.DEFAULT_FRAME_STYLE | wx.STAY_ON_TOP, title=u'Clipboard Monitor...')
        self.SetClientSize(wx.Size(300, 400))
        self.SetBackgroundColour(wx.Colour(0, 128, 0))

        self.listBox1 = wx.ListBox(choices=[], id=wxID_FRAME1LB_CBLIST, name='listBox1', parent=self, pos=wx.Point(5, 5), size=wx.Size(290, 350), style=0)
        self.listBox1.SetBackgroundColour(wx.Colour(255, 255, 128))
        self.listBox1.Bind(wx.EVT_LISTBOX, self.OnListBox1Listbox,
        id=wxID_FRAME1LB_CBLIST)

        self.button1 = wx.Button(id=wxID_FRAME1BTN_CLEAR, label=u'Stop Speak!', name='button1', parent=self, pos=wx.Point(5, 365),
        size=wx.Size(140, 30), style=0)
        self.button1.Bind(wx.EVT_BUTTON, self.OnBtnClearItem, id=wxID_FRAME1BTN_CLEAR)  

        self.button2 = wx.Button(id=wxID_FRAME1BTN_CLEARALL, label=u'Clear all', name='button2', parent=self, pos=wx.Point(155, 365),
        size=wx.Size(140, 30), style=0)
        self.button2.Bind(wx.EVT_BUTTON, self.OnBtnClearAll, id=wxID_FRAME1BTN_CLEARALL)            
               
        self.ignoreNotify = False
        
    def __init__ (self, parent):
        self._init_ctrls(parent)
        
        
        
        self.first   = True
        self.nextWnd = None

        # Get native window handle of this wxWidget Frame.
        self.hwnd    = self.GetHandle ()

        # Set the WndProc to our function.
        self.oldWndProc = win32gui.SetWindowLong (self.hwnd,
                                                  win32con.GWL_WNDPROC,
                                                  self.MyWndProc)

        try:
            self.nextWnd = win32clipboard.SetClipboardViewer (self.hwnd)
        except win32api.error:
            if win32api.GetLastError () == 0:
                # information that there is no other window in chain
                pass
            else:
                raise

    def MyWndProc (self, hWnd, msg, wParam, lParam):
        if msg == win32con.WM_CHANGECBCHAIN:
            self.OnChangeCBChain (msg, wParam, lParam)
        elif msg == win32con.WM_DRAWCLIPBOARD:
            self.OnDrawClipboard (msg, wParam, lParam)

        # Restore the old WndProc. Notice the use of win32api
        # instead of win32gui here. This is to avoid an error due to
        # not passing a callable object.
        if msg == win32con.WM_DESTROY:
            if self.nextWnd:
               win32clipboard.ChangeClipboardChain (self.hwnd, self.nextWnd)
            else:
               win32clipboard.ChangeClipboardChain (self.hwnd, 0)

            win32api.SetWindowLong (self.hwnd,
                                    win32con.GWL_WNDPROC,
                                    self.oldWndProc)

        # Pass all messages (in this case, yours may be different) on
        # to the original WndProc
        return win32gui.CallWindowProc (self.oldWndProc,
                                        hWnd, msg, wParam, lParam)

    def OnChangeCBChain (self, msg, wParam, lParam):
        if self.nextWnd == wParam:
           # repair the chain
           self.nextWnd = lParam
        if self.nextWnd:
           # pass the message to the next window in chain
           win32api.SendMessage (self.nextWnd, msg, wParam, lParam)

    def OnDrawClipboard (self, msg, wParam, lParam):
        if self.first:
           self.first = False
        else:
            if not self.ignoreNotify :
                win32clipboard.OpenClipboard()
                if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_TEXT):
                    data = win32clipboard.GetClipboardData()
                    self.listBox1.Append( data )
                    sentence = self.listBox1.GetString( self.listBox1.GetCount()-1 )
                    Engine.Speak( sentence, Flags=3 )
                                           
                    self.SetTitle(data)     
                else:
                    self.SetTitle("Clipboard monitor...")
                win32clipboard.CloseClipboard()

        self.ignoreNotify = False
            
        if self.nextWnd:
           # pass the message to the next window in chain
           win32api.SendMessage (self.nextWnd, msg, wParam, lParam)

    def OnListBox1Listbox(self, event):
        '''
        click list item and display the selected string in frame's title
        '''
        selName = self.listBox1.GetStringSelection()
        self.SetTitle(selName)
        self.ignoreNotify = True
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(selName)
        
        Engine.Speak( selName, Flags=3 )

        win32clipboard.CloseClipboard()   
    def OnBtnClearAll(self, event):
        self.listBox1.Clear()           
        
    def OnBtnClearItem(self, event):
        Engine.Speak( "", Flags=3 )
        

app   = wx.PySimpleApp ()
frame = create(None)
frame.Show ()
app.MainLoop ()
