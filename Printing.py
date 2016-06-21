import wx, os
FONTSIZE = 14

class TextDocPrintout( wx.Printout ):

    def __init__( self, text, title, margins ):
        wx.Printout.__init__( self, title )
        self.lines = text.split( '\n' )
        self.margins = margins
        self.numPages = 1

    def HasPage( self, page ):
        return page <= self.numPages

    def GetPageInfo( self ):
       return ( 1, self.numPages, 1, self.numPages )

    def CalculateScale( self, dc ):
        ppiPrinterX, ppiPrinterY = self.GetPPIPrinter()
        ppiScreenX, ppiScreenY = self.GetPPIScreen()
        logScale = float( ppiPrinterX )/float( ppiScreenX )
        pw, ph = self.GetPageSizePixels()
        dw, dh = dc.GetSize()
        scale = logScale * float( dw )/float( pw )
        dc.SetUserScale( scale, scale )
        self.logUnitsMM = float( ppiPrinterX )/( logScale*25.4 )

    def CalculateLayout( self, dc ):
        topLeft, bottomRight = self.margins
        dw, dh = dc.GetSize()
        self.x1 = topLeft.x * self.logUnitsMM
        self.y1 = topLeft.y * self.logUnitsMM
        self.x2 = ( dc.DeviceToLogicalXRel( dw) - bottomRight.x * self.logUnitsMM )
        self.y2 = ( dc.DeviceToLogicalYRel( dh ) - bottomRight.y * self.logUnitsMM )
        self.pageHeight = self.y2 - self.y1 - 2*self.logUnitsMM
        font = wx.Font( FONTSIZE, wx.TELETYPE, wx.NORMAL, wx.NORMAL )
        dc.SetFont( font )
        self.lineHeight = dc.GetCharHeight()
        self.linesPerPage = int( self.pageHeight/self.lineHeight )

    def OnPreparePrinting( self ):
        dc = self.GetDC()
        self.CalculateScale( dc )
        self.CalculateLayout( dc )
        self.numPages = len(self.lines) / self.linesPerPage
        if len(self.lines) % self.linesPerPage != 0:
            self.numPages += 1

    def OnPrintPage(self, page):
        dc = self.GetDC()
        self.CalculateScale(dc)
        self.CalculateLayout(dc)
        dc.SetPen(wx.Pen("black", 0))
        dc.SetBrush(wx.TRANSPARENT_BRUSH)
        r = wx.RectPP((self.x1, self.y1), (self.x2, self.y2))
        dc.DrawRectangleRect(r)
        dc.SetClippingRect(r)
        line = (page-1) * self.linesPerPage
        x = self.x1 + self.logUnitsMM
        y = self.y1 + self.logUnitsMM
        while line < (page * self.linesPerPage):
            dc.DrawText(self.lines[line], x, y)
            y += self.lineHeight
            line += 1
            if line >= len(self.lines):
                break
        return True

