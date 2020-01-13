def init():
    global img_dir
    global http_head
    global tmp_dir
    global mathtype_convert_to
    global  mode_text
    global subject
    img_dir=''
    http_head=''
    tmp_dir=''
    mathtype_convert_to='png'   ##png or mathml
    mode_text = r'(\d{1,2})[～\-~](\d{1,2})[小]{0,1}题'
    subject=''