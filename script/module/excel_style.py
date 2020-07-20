from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

style = {
    # タイトル
    'style_title' : [
        Font(
            bold = True,
            underline='double'
        ),
        PatternFill(
            
        ),
        Border(
            
        ),
        Alignment(
            
        )
    ],
    
    # ヘッダー
    'style_header' : [
        Font(
            bold = True,
            underline='none',
            color = '000000'
        ),
        PatternFill(
            'solid',
            fgColor='4682b4'
        ),
        Border(
            left=Side(border_style='thin',
                       color='000000'),
            right=Side(border_style='thin',
                       color='000000'),
            top=Side(border_style='thin',
                       color='000000'),
            bottom=Side(border_style='thin',
                       color='000000')
        ),
        Alignment(horizontal = 'center', 
                  vertical = 'center',
                  wrap_text = False
        )
    ],
    # 表
    'style_grid' : [
        Font(
            bold = False,
            underline='none',
            color = '000000'
        ),
        PatternFill(
            'solid',
            fgColor='ffffff'
        ),
        Border(
            left=Side(border_style='thin',
                       color='000000'),
            right=Side(border_style='thin',
                       color='000000'),
            top=Side(border_style='thin',
                       color='000000'),
            bottom=Side(border_style='thin',
                       color='000000')
        ),
        Alignment(horizontal = 'center', 
                  vertical = 'center',
                  wrap_text = False
        )
    ]   
}