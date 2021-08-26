import streamlit as st
import pandas as pd
import yagmail


# 標題
st.markdown("<h1 style='text-align: center; color: blue;'>Mail Sender</h1>", unsafe_allow_html=True)
st.markdown('---')

###### 寄件者人 Gmail 帳密 ######
st.sidebar.header('寄件人資訊')
account = st.sidebar.text_input('帳號')
password = st.sidebar.text_input('密碼', type='password')

###### 選擇寄送項目 ########
st.header('選擇寄送項目')
option = st.selectbox('', options=['收據', '連結'])


###### 上傳參加者資訊 ######
if option == '收據':
    st.header('參加者資訊表')
    participant_info = st.file_uploader('', type=['csv'])
    if participant_info is not None:
        try:
            participant_info_df = pd.read_csv(participant_info)
        except:
            participant_info_df = pd.read_csv(participant_info, encoding = 'ansi')

        participant_info_df['學號'] = participant_info_df['學號'].apply(lambda x : x.upper())
        check_info = st.checkbox('展開資料表')
        if check_info:
            st.write(participant_info_df.to_html(escape = False), unsafe_allow_html = True)


    ###### 收據資料上傳 #######
    st.header('上傳收據資料')
    receipts = st.file_uploader('', accept_multiple_files=True, type=['docx'])

    ##### 信件內容 #######
    st.header('信件內容')
    mail_content = st.text_area('', height=300, max_chars=None, key=None)

    ##### 確認是否設定完成 ######
    set_up = st.button('開始寄信')

    ##### 開始寄信 ######
    
    send_finish = False

    if set_up == True:
        yag = yagmail.SMTP(account, password)

        status_list = []

        for receipt in receipts:
            # 配對收據檔案
            studentID = receipt.name.strip('.docx')
            info = participant_info_df.loc[(participant_info_df['學號'] == studentID)]
            # 基本資料
            for idx, row in info.iterrows():
                name = row['姓名']
                receiver_email = row['email']
                content = name + mail_content
            # 寄信
            try:    
                yag.send(receiver_email, 
                        subject = 'TASSEL_VC_實驗收據', 
                        contents = content,
                        attachments = receipt)

                success = 'successful'

            except:
                success = 'failed'
            
            status_list.append({'姓名':name,'信箱':receiver_email, '狀態':success})
            send_finish = True

    ###### 檢視寄信狀態 ######
    st.markdown('---')
    st.header('寄件狀態')
    if send_finish:
        status_df = pd.DataFrame(status_list)
        status_df.sort_values(['狀態'], inplace = True)
        st.write(status_df.to_html(escape = False), unsafe_allow_html = True)


else: 
    # 連結資料表
    st.header('連結資訊表') 
    link_info = st.file_uploader('', type=['csv']) 
    if link_info is not None:
        try:
            link_info_df = pd.read_csv(link_info)
        except:
            link_info_df = pd.read_csv(link_info, encoding = 'ansi')

        link_info_df['學號'] = link_info_df['學號'].apply(lambda x : x.upper())
        check_info = st.checkbox('展開資料表')
        if check_info:
            st.write(link_info_df.to_html(escape = False), unsafe_allow_html = True)

    # 信件內容
    st.header('信件內容')
    mail_content = st.text_area('連結部分寫上\"連結\"', height=300, max_chars=None, key=None)
    
    # 是否設定完成
    set_up = st.button('開始寄信')

    ##### 開始寄信 ######

    send_finish = False

    if set_up == True:

        yag = yagmail.SMTP(account, password)

        status_list = []

        for idx, row in link_info_df.iterrows():
            receiver_email = row['email'] 
            link = row['連結']

            content = mail_content.replace('連結', link)

            try:
                yag.send(receiver_email, 
                        subject = 'TASSEL_VC_實驗連結', 
                        contents = content)
                success = 'successful'
            except:
                success = 'failed'

            status_list.append({'姓名':name,'信箱':receiver_email, '狀態':success})
            send_finish = True

    ##### 檢視寄信狀態 ######
    st.markdown('---')

    st.header('寄件狀態')
    if send_finish:
        status_df = pd.DataFrame(status_list)
        status_df.sort_values(['狀態'], inplace = True)
        st.write(status_df.to_html(escape = False), unsafe_allow_html = True)