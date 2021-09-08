import streamlit as st
import pandas as pd 
import yagmail
import docx2txt


######## Title ########
st.markdown("<h1 style='text-align: center;'>Mail Sender</h1>", unsafe_allow_html=True)
option = st.selectbox('寄送項目', options=['收據', '實驗連結'])
st.markdown('---')

######## Sender information ########
st.sidebar.header('寄件人資訊')

account = st.sidebar.text_input('Google 信箱')
password = st.sidebar.text_input('應用程式密碼', type='password')
st.sidebar.markdown('[帳號設定說明](https://docs.google.com/document/d/1YQMeRTibK7qjZonhz6Xl3a0Jv3-rj_g9o7dyvOVTxBk/edit?usp=sharing)')

######## funtion ########
def modify_df(df = None):
    """- convert all columns to string  
       - convert all lowercase letters in studentID to uppercase letters 
    """
    df = df.astype(str)
    df['學號'] = df['學號'].apply(lambda x : x.upper())

    return df

######## Receipts ########
if (option == '收據'):

    # Participants information

    st.header('參加者資料表')
    st.markdown('- 檔案須為 CSV UTF-8 檔')
    st.markdown('- 須包含的欄位有學號、姓名、email (欄位名稱請一致)') 
    st.markdown('- 學號中的英文字母將全數轉換為大寫')

    participant_info = st.file_uploader('', type = ['csv'])

    if participant_info != None:
        participant_info_df = pd.read_csv(participant_info)
        participant_info_df = modify_df(participant_info_df)
        check_info = st.checkbox('展開資料表')
    
        if check_info:
            st.write(participant_info_df.to_html(escape = False), unsafe_allow_html = True)

    # Upload receipts
    st.header('上傳收據資料')
    st.markdown('- 檔案為 PDF 檔')
    st.markdown('- 檔案以學號命名, 且英文字母一律大寫')

    receipts = st.file_uploader('', accept_multiple_files=True, type=['pdf'])

    # Mail content
    st.header('信件內容')
    st.markdown('- 信件開頭會自動加上同學姓名')
    receipt_subject = st.text_input('信件主旨', value='TASSEL_VC 實驗收據')
    script = docx2txt.process("receipt_content.docx")
    receipt_content = st.text_area('信件內容', value =script, height=300, max_chars=None, key=None)

    # Finish setting up
    set_up = st.button('開始寄信')
    st.markdown('---')

    send_finish = False

    # Sending receipts
    if (set_up == True):
        
        yag = yagmail.SMTP(account, password)

        status_list = []
        no_match_receipt = []
        
        for receipt in receipts:
            studentID = receipt.name.strip('.pdf') 
            info = participant_info_df.loc[(participant_info_df['學號'] == studentID)] # match receipt to student

            if len(info) == 0:
                no_match_receipt.append(studentID + '.pdf')
            else: 
                name = info['姓名'].values[0]
                receiver_email = info['email'].values[0]
                content = name + receipt_content

                try:
                    yag.send(to = receiver_email, subject = receipt_subject, contents = receipt_content, attachments = receipt)
                    status = 'successful'
                except:
                    status = 'failed'

                status_list.append({'姓名':name, 'email':receiver_email, '狀態':status})

        send_finish = True

    # check for send mail status
    st.header('寄信狀態檢視')
    if send_finish == True:
        
        st.subheader('寄件狀態')
        status_df = pd.DataFrame(status_list)
        st.write(status_df.to_html(escape = False), unsafe_allow_html = True)  

        st.subheader('缺少收據的名單')
        st.markdown('- 被列出者代表沒有他的收據')
        no_receive = set(participant_info_df['姓名']) - set(status_df['姓名'])
        st.write(no_receive)

        st.subheader('找不到主人的收據')
        st.markdown('- 未能配對參加者的收據檔案')
        st.write(set(no_match_receipt))

######## Link ########
else:
    # link information
    st.header('實驗連結資訊表') 
    st.markdown('- 檔案為 CSV UTF-8 檔')
    st.markdown('- 欄位須包含學號、連結')

    link_info = st.file_uploader('', type=['csv']) 

    if link_info is not None:
        link_info_df = pd.read_csv(link_info)
        link_info_df = modify_df(link_info_df)
        check_info = st.checkbox('展開資料表')
    
        if check_info:
            st.write(link_info_df.to_html(escape = False), unsafe_allow_html = True)

    # Content
    st.header('信件內容')
    st.markdown('- 將要傳送的連結, 以\"連結\"文字代替 (不需要引號)')
        
    link_subject = st.text_input('信件主旨', value = 'TASSEL_VC 實驗連結')
    script = docx2txt.process("link_content.docx")
    link_content = st.text_area('信件內容', value =script, height=300, max_chars=None, key=None)


    # Finish setting up
    set_up = st.button('開始寄信')
    st.markdown('---')

    send_finish = False

    # Send mail
    if set_up == True:

        status_list = []

        yag = yagmail.SMTP(account, password)

        for idx, row in link_info_df.iterrows():
            name = row['姓名']
            receiver_email = row['email']
            link = row['連結']
            content = link_content.replace('連結', link)

            try:
                yag.send(to = receiver_email, subject = link_subject, contents = content)
                status = 'successful'
            except:
                status = 'failed'

            status_list.append({'姓名':name, 'email':receiver_email, '狀態':status})

        send_finish = True

    # check for send mail status
    st.header('寄信狀態檢視')
    if send_finish == True:        
        st.subheader('寄件狀態')
        status_df = pd.DataFrame(status_list)
        status_df.sort_values(['狀態'], inplace = True)
        st.write(status_df.to_html(escape = False), unsafe_allow_html = True)  
