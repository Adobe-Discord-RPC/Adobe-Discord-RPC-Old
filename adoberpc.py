# -*- coding:utf-8 -*- 

"""
                  _       _            _____  _                       _   _____  _____   _____ 
         /\      | |     | |          |  __ \(_)                     | | |  __ \|  __ \ / ____|
        /  \   __| | ___ | |__   ___  | |  | |_ ___  ___ ___  _ __ __| | | |__) | |__) | |     
       / /\ \ / _` |/ _ \| '_ \ / _ \ | |  | | / __|/ __/ _ \| '__/ _` | |  _  /|  ___/| |     
      / ____ \ (_| | (_) | |_) |  __/ | |__| | \__ \ (_| (_) | | | (_| | | | \ \| |    | |____ 
     /_/    \_\__,_|\___/|_.__/ \___| |_____/|_|___/\___\___/|_|  \__,_| |_|  \_\_|     \_____|
 
    Ver. 3.5 (Hotfix)
    © 2017-2020 Adobe Discord RPC Team.
    Follows GPL-3.0
    Gtihub || https://github.com/Adobe-Discord-RPC

    :: Program Core ::
"""

if __name__ == "__main__" :
    # 20-11-02 임시로 언어번역 추가 :: v4.0에 해당 번역 이전 계획 없음

    def log(tpe, inf, datetime = None):
        # 20-11-01 로그파일이 너무 방대해져서 실행 시마다 기존파일 삭제로 변경
        # 20-11-01 디버그 조건 변경예정
        if (tpe == "INFO") and (inf == "Init"): # && << 이거 되는거 아니었어???????????????????????????????????????????
            if os.path.isfile('./log.log'):
                os.remove('./log.log')
            open('./log.log', 'w').close()
        if datetime == None:
            prnt = "Time UNK | %s | %s" % (tpe, inf)
        else:
            now = datetime.datetime.now()
            prnt = "%s-%02d-%02d %02d:%02d:%02d | %s | %s" % (now.year, now.month, now.day, now.hour, now.minute, now.second, tpe, inf)
        f = open("./log.log", 'a', encoding='utf8')
        f.write(prnt+"\n")
        f.close()
        print(prnt)

    def goout(datetime = None):
        log("INFO", "Ext", datetime)
        #time.sleep(30)
        #os.system('adoberpc.exe')
        exit()
        #sys.exit()

    try:
        from pypresence import Presence
        import datetime, os, requests, sys, psutil, json, win32gui, win32process, time, pandas, plyer, re, win32ui, win32con, platform, locale
    except ModuleNotFoundError as e:
        if not (str(e) in 'plyer'):
            plyer.notification.notify(
                title='Adobe Discord RPC',
                message='모듈 불러오기에 실패했습니다.\nModule fetch failed.',
                app_name='Adobe Discord RPC',
                app_icon='icon_alpha.ico'
            )
        log("ERROR", "Module NF : %s" % (str(e).replace('No module named ', '')))
        goout()

    # Init
    log("INFO", "Init")
    dur = False

    # 언어 Detect
    if locale.getdefaultlocale() == "ko_KR": # 한어
        langcode = "Ko"
        lang = {
            'Current': '(현재)', 'Latest': '(최신)',
            'NewVerNoti': 'Adobe Discord RPC가 가동되었습니다.\n새 버전이 있습니다.',
            'NewVerDial': '새 버전이 있습니다.\n업데이트를 진행할까요?',
            'DiscordFailed': '디스코드와 연결하지 못하였습니다.\n디스코드가 켜져 있는지 다시 한번 확인 해 주세요.\n30초 후 연결을 다시 시도합니다.',
            'Connected01': '성공적으로 디스코드와 연결했습니다.', 'Connected02': '(을)를 플레이 하게 됩니다.',
            'UpdateFailed': 'RPC 갱신에 실패하였습니다.\n30초 후 다시 시도합니다.'
        }
    else: # 영어로 실행
        langcode = "En"
        lang = {
            'Current': '(Current)', 'Latest': '(Latest)',
            'NewVerNoti': 'Adobe Discord RPC ran successfully.\nYou have a new version.',
            'NewVerDial': 'You have a new version.\nupdate?',
            'DiscordFailed': 'Failed to connect to Discord.\nPlease check again to see if the Discord is on.\nRetry the connection after 30 seconds.',
            'Connected01': 'Successfully connected with Discord.', 'Connected02': 'Will be shown on your profile.',
            'UpdateFailed': 'Failed to update RPC.\nAutomatically retry after 30 seconds.'
        }

    log("DEBUG", "langcode : %s, %s" % (langcode, lang))
        
    try:
        os.remove('./stop.req')
    except FileNotFoundError:
        pass

    def checkver():
        # 20-03-23 방식 변경
        #          새 버전 알림 -> 새 버전 알림 후 업데이트 수락하면 업데이터 가동
        with open('programver.json', encoding='utf8') as f:
            data = json.load(f)

        # 20-11-01 v4.0 개발로 인한 백업서버 미사용 (서버 구조 변경예정임)
        # 20-11-01 기존 서버는 v4.0 이후 비활성 계획중임
        nowver = data['ver']
        r = requests.get("https://cdn.adoberpc.hwahyang.space/adoberpc_ver.json")
        if r.status_code != 200:
            log("DEBUG", "서버 연결 실패.", datetime)
            r = "{\"ver\": 0}"
        else:
            r = r.text
        data = json.loads(r)
        latest = float(data["ver"]) # 바보야 int 아니라고,,
        if latest > nowver: # 만약에 최신이 더 높다면,
            log("DEBUG", "새 버전 알림 발신. || 현재 : %s || 최신 : %s" % (nowver, latest), datetime)
            plyer.notification.notify(
                title='Adobe Discord RPC',
                # message='Adobe Discord RPC가 가동되었습니다.\n새 버전이 있습니다.\nV%s (현재) -> V%s (최신)' % (nowver, latest),
                message='%s\nV%s %s -> V%s %s' % (lang['NewVerNoti'], nowver, lang['Current'], latest, lang['Latest']),
                app_name='Adobe Discord RPC',
                app_icon='icon_alpha.ico'
            )
            #res = win32ui.MessageBox("새 버전이 있습니다.\nV%s (현재) -> V%s (최신)\n업데이트를 진행할까요?" % (nowver, latest), "Adobe Discord RPC", win32con.MB_YESNO)
            res = win32ui.MessageBox("%s\nV%s %s -> V%s %s" % (lang['NewVerDial'], nowver, lang['Current'], latest, lang['Latest']), "Adobe Discord RPC", win32con.MB_YESNO)
            if res == win32con.IDYES:
                os.system('start adoberpc_updater.exe')
                open("stop.req", 'w').close()
                goout()
        else:
            log("DEBUG", "버전 변동 없음. || 현재 : %s || 최신 : %s" % (nowver, latest), datetime)
            #notification.notify(
            #    title='Adobe Discord RPC',
            #    message='Adobe Discord RPC가 가동되었습니다.\n새 버전이 존재하지 않습니다.',
            #    app_name='Adobe Discord RPC',
            #    app_icon='icon_alpha.ico'
            #)

    with open('pinfo.json', encoding='utf8') as f:
        data = json.load(f)

    def get_title(pid):
        def callback(hwnd, hwnds):
            if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                if found_pid == pid:
                    hwnds.append(hwnd)
        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        window_title = win32gui.GetWindowText(hwnds[-1])
        return window_title

    def get_info(programname):
        getp = lambda process: (list(p.info for p in filter((lambda p: p.info['name'] and p.info['name'] == process),list(psutil.process_iter(['pid','name','exe','status'])))))
        returns = getp(programname)

        if not returns: # 미친 이거 왜됨
            return [False]
        
        returns = returns[0]

        try:
            windowname = get_title(returns['pid'])
        except Exception as e:
            print(e)
            return [False]

        return [True, returns['name'], returns['pid'], windowname, returns['status'], returns['exe']]

    def get_process_info():
        for now_json in data:
            def_returned = get_info(now_json['processName'])
            if def_returned[0]:
                return def_returned + [now_json]
            else:
                continue

        return [False]

    def get_window_title(pid):
        def callback(hwnd, hwnds):
            if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                if found_pid == pid:
                    hwnds.append(hwnd)
        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        window_title = win32gui.GetWindowText(hwnds[-1])
        return window_title

    if not (sys.platform in ['Windows', 'win32', 'cygwin']):
        # 않이 근데 왠만해선 타 OS에서 돌릴 일은 없지 않나
        log("ERROR", "Unsupported OS : %s" % (sys.platform))
        goout(datetime)

    if int(platform.release()) < 7:
        log("ERROR", "Unsupported Windows Version : %s" % (platform.release()))
        goout(datetime)

    checkver()

    def do():
        try: # 애러나면 로그는 해야죠
            while True:
                # 죽을때까지 실행 검증 루프
                a = True
                while a: # 실행 없으면 15초 쉬었다 루프 다시. --> 30초로 변경.
                    try:
                        pinfo = get_process_info()
                        if not pinfo[0]:
                            log("DEBUG", "Process NF", datetime)
                            log("DEBUG", "Retry in 30s...", datetime)
                            time.sleep(30)
                            continue
                        window_title = pinfo[3]
                        a = False # 루프 종료
                    except IndexError as e:
                        log("ERROR", "caught error.. Retry in 10s : %s" % (e), datetime)
                        time.sleep(10)

                RPC = Presence(pinfo[len(pinfo)-1]['appid'])

                try:
                    RPC.connect()
                except Exception as e:
                    log("ERROR", "Discord Connection Failed", datetime)
                    log("DEBUG", "ERRINFO :: %s" % (e), datetime)
                    plyer.notification.notify(
                        title='Adobe Discord RPC',
                        #message='디스코드와 연결하지 못하였습니다.\n디스코드가 켜져 있는지 다시 한번 확인 해 주세요.\n30초 후 연결을 다시 시도합니다.',
                        message=lang['DiscordFailed'],
                        app_name='Adobe Discord RPC',
                        app_icon='icon_alpha.ico'
                    )
                    goout(datetime)
                log("INFO", "Discord Connect Success", datetime)

                dt = pandas.to_datetime(datetime.datetime.now())
                # -1은 미사용
                # 그 이상은 splitindex 용도로 사용함
                if pinfo[len(pinfo)-1]['getver'] == -1:
                    version = ''
                else:
                    path = pinfo[5].split(pinfo[len(pinfo)-1]['publicName'])[1]
                    path = path.split('\\')[pinfo[len(pinfo)-1]['getver']]
                    version = path.replace(" ", '')

                filename = window_title.split(pinfo[len(pinfo)-1]['splitBy'])[pinfo[len(pinfo)-1]['splitIndex']]
                lmt = pinfo[len(pinfo)-1][langcode]['largeText'].replace('%Ver%', version)
                lmt = lmt.replace('%Filename%', filename)
                smt = pinfo[len(pinfo)-1][langcode]['smallText'].replace('%Filename%', filename)
                smt = smt.replace('%Ver%', version)
                b = True

                plyer.notification.notify(
                    title='Adobe Discord RPC',
                    #message='성공적으로 디스코드와 연결했습니다.\n%s를 플레이 하게 됩니다.' % (pinfo[len(pinfo)-1]['publicName']),
                    message='%s\n%s %s' % (lang['Connected01'], pinfo[len(pinfo)-1]['publicName'], lang['Connected02']),
                    app_name='Adobe Discord RPC',
                    app_icon='icon_alpha.ico'
                )

                while b:
                    # 20-04-25 코드 구조 대량으로 변경되어서 응답없음 여부 사용 안함
                    """if not isresponding(pinfo['processName'].replace('.exe', ''), datetime):
                        # 응답없음은 10초 간격으로 체크해서 응답없음 풀리면 c = False 로 만듬
                        c = True
                        while c:
                            try:
                                rtn = RPC.update(
                                    large_image='lg', large_text='프로그램 : %s' % (pinfo['publicName']),
                                    small_image='sm_temp', small_text="파일명 : %s" %(filename),
                                    details="응답없음", state="응답없음",
                                    start=int(time.mktime(dt.timetuple()))
                                )
                            except Exception as e:
                                log("ERROR", "pypresence 갱신 실패... 10초 대기 : %s" % (e), datetime)
                                plyer.notification.notify(
                                    title='Adobe Discord RPC',
                                    message='RPC 갱신에 실패하였습니다.\n10초 후 다시 시도합니다.',
                                    app_name='Adobe Discord RPC',
                                    app_icon='icon_alpha.ico'
                                )
                                b = False
                                time.sleep(10)
                            else:
                                log("DEBUG", "pypresence 리턴 : %s" % (rtn), datetime)
                                time.sleep(10)
    
                                if not isresponding(pinfo['processName'].replace('.exe', ''), datetime):
                                    pass
                                else:
                                    # 정상실행 경우
                                    # RPC.clear(pid=os.getpid())
                                    c = False
                                    pass
                    else:"""
                    try:
                        rtn = RPC.update(
                            large_image='lg', large_text='Program : %s' % (pinfo[len(pinfo)-1]['publicName']),
                            small_image='sm_temp', small_text="Adobe Discord RPC",
                            details=lmt, state=smt,
                            start=int(time.mktime(dt.timetuple()))
                        )
                    except Exception as e:
                        log("ERROR", "pypresence 갱신 실패... 30초 대기 : %s" % (e), datetime)
                        plyer.notification.notify(
                            title='Adobe Discord RPC',
                            #message='RPC 갱신에 실패하였습니다.\n30초 후 다시 시도합니다.',
                            message=lang['UpdateFailed'],
                            app_name='Adobe Discord RPC',
                            app_icon='icon_alpha.ico'
                        )
                        b = False
                        time.sleep(30)
                    else:
                        log("DEBUG", "pypresence Return : %s" % (rtn), datetime)
                        time.sleep(30)
                        #time.sleep(5) # 테스트용인데 왜 주석 안하고 배포한거야

                        dur = True
                        window_title = get_window_title(pinfo[2])
                        filename = window_title.split(pinfo[len(pinfo)-1]['splitBy'])[pinfo[len(pinfo)-1]['splitIndex']]
                        lmt = pinfo[len(pinfo)-1]['largeText'].replace('%Ver%', version)
                        lmt = lmt.replace('%Filename%', filename)
                        smt = pinfo[len(pinfo)-1]['smallText'].replace('%Filename%', filename)
                        smt = smt.replace('%Ver%', version)
                        dur = False
                        pass

        except KeyboardInterrupt:
            log("DEBUG", "KeyboardInterrupted by console", datetime)
            goout()
        except Exception as e:
            if dur:
                # PID 변경 OR 프로그램 종료. 같은 PID 금방 다시 할당할 일 없음.
                log("DEBUG", "PID 변동 OR 프로세스 종료. : %s -> %s" % (pinfo, e), datetime)
                RPC.clear(pid=os.getpid())
                b = False
                do()
                pass
            else:
                log("ERROR", "Undefined : %s" % (e), datetime)
                goout()

    do()

# End of Code.