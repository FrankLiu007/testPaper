def init():
    global img_dir
    global http_head
    global tmp_dir
    global mathtype_convert_to
    global  mode_text
    global subject
    global strict_mode
    global dati_mode
    global material_mode
    global curr_ti_number
    global jie_mode
    global xiaoti_mode
    img_dir=''
    http_head=''
    tmp_dir=''
    mathtype_convert_to='png'   ##png or mathml
    mode_text = r'(\d{1,2})[～到至、\-~](\d{1,2})[小]{0,1}题'
    material_mode=r'(\d{1,2})[～到至、\-~](\d{1,2})[小]{0,1}题'
    subject=''
    strict_mode=True
    dati_mode=''
    jie_mode=''
    xiaoti_mode=r'^(\d{1,2})[\s\.．、]'
    curr_ti_number=-1

