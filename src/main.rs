#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use anyhow::{ensure, Context, Result};
use std::char::{decode_utf16, REPLACEMENT_CHARACTER};
use std::mem;
use std::path::PathBuf;
use std::slice;
use std::sync::{
    mpsc::{self, Sender},
    Mutex, OnceLock,
};
use std::thread;
use windows::{
    core::{w, Interface, HSTRING, PCWSTR, PWSTR},
    Foundation::TypedEventHandler,
    Media::{
        Core::MediaSource,
        Playback::MediaPlayer,
        SpeechSynthesis::{SpeechSynthesisStream, SpeechSynthesizer, VoiceInformation},
    },
    Storage::Streams::DataReader,
    Win32::{
        Foundation::{HWND, LPARAM, LRESULT, RECT, WPARAM},
        Graphics::Gdi::{
            BeginPaint, EndPaint, GetSysColorBrush, SetBkMode, TextOutW, UpdateWindow,
            COLOR_MENUBAR, PAINTSTRUCT, TRANSPARENT,
        },
        System::{LibraryLoader::GetModuleHandleW, WinRT::IBufferByteAccess},
        UI::{
            Controls::{
                Dialogs::{GetSaveFileNameW, OPENFILENAMEW},
                InitCommonControlsEx, ICC_BAR_CLASSES, INITCOMMONCONTROLSEX, TBM_SETPAGESIZE,
                TBM_SETPOS, TBM_SETRANGE, TBM_SETTICFREQ, TBS_AUTOTICKS, TBS_TOOLTIPS,
                WC_COMBOBOXW,
            },
            WindowsAndMessaging::{
                CreateWindowExW, DefWindowProcW, DispatchMessageW, GetClientRect, GetMessageW,
                GetWindowTextLengthW, GetWindowTextW, MessageBoxW, PostQuitMessage, RegisterClassW,
                SendMessageW, ShowWindow, TranslateMessage, BS_PUSHBUTTON, CBS_DROPDOWNLIST,
                CBS_HASSTRINGS, CBS_SORT, CB_ADDSTRING, CB_GETCURSEL, CB_GETLBTEXT,
                CB_SELECTSTRING, CW_USEDEFAULT, ES_AUTOVSCROLL, ES_MULTILINE, ES_WANTRETURN, HMENU,
                MB_OK, MSG, SW_SHOW, WINDOW_EX_STYLE, WINDOW_STYLE, WM_COMMAND, WM_CREATE,
                WM_DESTROY, WM_PAINT, WM_SETTEXT, WNDCLASSW, WS_BORDER, WS_CAPTION, WS_CHILD,
                WS_EX_STATICEDGE, WS_MINIMIZEBOX, WS_OVERLAPPED, WS_SYSMENU, WS_TABSTOP,
                WS_VISIBLE, WS_VSCROLL,
            },
        },
    },
};

/// メインウィンドウのクラス名
const CLASS_NAME: PCWSTR = w!("speech_window_cls42");
/// 再生ボタンの ID
const ID_PLAY: u16 = 5890;
/// クリアボタンの ID
const ID_CLEAR: u16 = 5891;
/// 保存ボタンの ID
const ID_SAVE: u16 = 5892;
/// コンボボックスの ID
const ID_COMBO: u16 = 5893;
/// トラックバーの ID
const ID_TRACKBAR: u16 = 5894;
/// エディットコントロールの [HWND](https://microsoft.github.io/windows-docs-rs/doc/windows/Win32/Foundation/struct.HWND.html) を保持するためのグローバル変数
static EDIT_HWND: OnceLock<Hwnd> = OnceLock::new();
/// コンボボックスの [HWND](https://microsoft.github.io/windows-docs-rs/doc/windows/Win32/Foundation/struct.HWND.html) を保持するためのグローバル変数
static COMBOBOX_HWND: OnceLock<Hwnd> = OnceLock::new();
/// トラックバーの [HWND](https://microsoft.github.io/windows-docs-rs/doc/windows/Win32/Foundation/struct.HWND.html) を保持するためのグローバル変数
static TRACKBAR_HWND: OnceLock<Hwnd> = OnceLock::new();
/// スピーチ再生スレッド実行待ちのための [Sender] を保持しておくグローバル変数
static STOP: Mutex<Vec<Sender<()>>> = Mutex::new(vec![]);

/// [HWND](https://microsoft.github.io/windows-docs-rs/doc/windows/Win32/Foundation/struct.HWND.html) をグローバル変数に保持するためのラッパ構造体
struct Hwnd(HWND);

/// [Hwnd] 構造体を別スレッドに送れるようにマーカトレイトである Send, Sync を実装する
unsafe impl Sync for Hwnd {}
unsafe impl Send for Hwnd {}

impl Hwnd {
    fn new(hwnd: HWND) -> Self {
        Self(hwnd)
    }

    fn handle(&self) -> HWND {
        self.0
    }
}

fn get_selected_voice_information() -> Result<VoiceInformation> {
    let hwnd = COMBOBOX_HWND.get().context("no handle")?.handle();
    let ret = unsafe { SendMessageW(hwnd, CB_GETCURSEL, None, None) };
    ensure!(ret.0 >= 0, "failed to get selected item index.");

    let buf = [0u16; 64];
    let ret = unsafe {
        SendMessageW(
            hwnd,
            CB_GETLBTEXT,
            WPARAM(ret.0 as _),
            LPARAM(buf.as_ptr() as _),
        )
    };

    SpeechSynthesizer::AllVoices()?
        .into_iter()
        .filter_map(|v| {
            if v.DisplayName().ok()?.as_wide() == &buf[..ret.0 as _] {
                Some(v)
            } else {
                None
            }
        })
        .next()
        .context("no voice.")
}

fn get_speaking_rate() -> Result<f64> {
    let hwnd = TRACKBAR_HWND.get().context("no handle.")?.handle();
    let ret = unsafe { SendMessageW(hwnd, 1024, None, None) }.0 as f64 / 10.0;
    ensure!(0.5 <= ret && ret <= 2.5, "invalid speaking rate.");
    Ok(ret)
}

fn speech_synthesis_stream(source: &[u16]) -> Result<SpeechSynthesisStream> {
    let source = HSTRING::from_wide(source)?;
    let synth = SpeechSynthesizer::new()?;
    let voice = get_selected_voice_information()?;
    synth.SetVoice(&voice)?;
    let speaking_rate = get_speaking_rate()?;
    synth.Options()?.SetSpeakingRate(speaking_rate)?;
    let stream = synth.SynthesizeTextToStreamAsync(&source)?.get()?;
    Ok(stream)
}

fn speech() -> Result<()> {
    let text = get_edit_control_text()?;
    thread::spawn(move || -> Result<()> {
        let stream = speech_synthesis_stream(&text)?;
        let player = MediaPlayer::new()?;
        let media_source = MediaSource::CreateFromStream(&stream, &stream.ContentType()?)?;
        player.SetSource(&media_source)?;
        let (tx, rx) = mpsc::channel();
        {
            let mut stop = STOP.lock().unwrap();
            stop.push(tx.clone());
        }
        let tx_clone = tx.clone();
        let token_media_ended = player.MediaEnded(&TypedEventHandler::new(move |_, _| {
            tx_clone.send(()).ok();
            Ok(())
        }))?;
        let token_media_failed = player.MediaFailed(&TypedEventHandler::new(move |_, _| {
            tx.send(()).ok();
            Ok(())
        }))?;
        player.Play()?;
        rx.recv()?;
        player.Close()?;
        player.RemoveMediaEnded(token_media_ended)?;
        player.RemoveMediaFailed(token_media_failed)?;
        Ok(())
    });
    Ok(())
}

fn get_save_file_path(hwnd: HWND) -> Result<PathBuf> {
    let mut buf = "speech.wav"
        .encode_utf16()
        .chain([0; 502])
        .collect::<Vec<_>>();
    let mut filename = OPENFILENAMEW {
        lStructSize: mem::size_of::<OPENFILENAMEW>() as _,
        hwndOwner: hwnd,
        lpstrFile: PWSTR::from_raw(buf.as_mut_ptr()),
        lpstrFilter: w!("Wave File (.wav)\0*.wav\0\0"),
        lpstrDefExt: w!("wav"),
        nMaxFile: buf.len() as _,
        ..Default::default()
    };
    unsafe { GetSaveFileNameW(&mut filename).ok()? };
    let path: String = decode_utf16(buf.iter().take_while(|v| *v != &0).copied())
        .map(|r| r.unwrap_or(REPLACEMENT_CHARACTER))
        .collect();
    Ok(path.into())
}

fn save_to_wav(hwnd: HWND) -> Result<()> {
    let file_path = get_save_file_path(hwnd)?;

    let text = get_edit_control_text()?;
    let stream = speech_synthesis_stream(&text)?;
    let reader = DataReader::CreateDataReader(&stream)?;
    let size = stream.Size()? as u32;
    reader.LoadAsync(size)?.get()?;
    let buffer: IBufferByteAccess = reader.ReadBuffer(size)?.cast()?;
    let ptr = unsafe { buffer.Buffer()? };

    let slice = unsafe { slice::from_raw_parts(ptr, size as usize) };
    std::fs::write(&file_path, slice)?;

    let file_name = file_path.file_name().context("no file name.")?;
    let msg = format!("{} を保存しました。", file_name.to_string_lossy());
    let msg = msg.encode_utf16().chain(Some(0)).collect::<Vec<_>>();
    unsafe { MessageBoxW(hwnd, PCWSTR(msg.as_ptr()), w!("speech"), MB_OK) };
    Ok(())
}

fn paint(hwnd: HWND) -> Result<()> {
    let mut ps = PAINTSTRUCT::default();
    let hdc = unsafe { BeginPaint(hwnd, &mut ps) };
    unsafe { SetBkMode(hdc, TRANSPARENT) };
    unsafe { TextOutW(hdc, 10, 50, w!("読み上げ速度：遅").as_wide()).ok()? };
    unsafe { TextOutW(hdc, 550, 50, w!("速").as_wide()).ok()? };
    unsafe { EndPaint(hwnd, &mut ps).ok()? };
    Ok(())
}

fn get_edit_control_text() -> Result<Vec<u16>> {
    let hwnd = EDIT_HWND.get().context("no handle.")?.handle();
    let len = unsafe { GetWindowTextLengthW(hwnd) };
    let mut buf = vec![0; len as usize + 1];
    unsafe { GetWindowTextW(hwnd, &mut buf) };
    Ok(buf)
}

fn clear_edit_control_text() -> Result<()> {
    let hwnd = EDIT_HWND.get().context("no handle.")?.handle();
    unsafe { SendMessageW(hwnd, WM_SETTEXT, None, None) };
    let mut stop = STOP.lock().unwrap();
    while !stop.is_empty() {
        if let Some(tx) = stop.pop() {
            _ = tx.send(());
        }
    }
    Ok(())
}

fn command(hwnd: HWND, wparam: WPARAM) -> Result<()> {
    let id = loword(wparam.0 as _);

    if id.eq(&ID_PLAY) {
        speech()?;
    } else if id.eq(&ID_CLEAR) {
        clear_edit_control_text()?;
    } else if id.eq(&ID_SAVE) {
        save_to_wav(hwnd)?;
    }

    Ok(())
}

fn create_button(
    hwnd: HWND,
    label: PCWSTR,
    x: i32,
    y: i32,
    width: i32,
    height: i32,
    id: u16,
) -> Result<()> {
    unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("BUTTON"),
            label,
            WS_CHILD | WS_VISIBLE | WINDOW_STYLE(BS_PUSHBUTTON as _),
            x,
            y,
            width,
            height,
            hwnd,
            HMENU(id as _),
            None,
            None,
        )?
    };
    Ok(())
}

fn create_play_button(hwnd: HWND) -> Result<()> {
    create_button(hwnd, w!("再生"), 10, 10, 100, 30, ID_PLAY)?;
    Ok(())
}

fn create_clear_button(hwnd: HWND) -> Result<()> {
    create_button(hwnd, w!("クリア"), 120, 10, 100, 30, ID_CLEAR)?;
    Ok(())
}

fn create_save_button(hwnd: HWND) -> Result<()> {
    create_button(hwnd, w!("保存"), 230, 10, 100, 30, ID_SAVE)?;
    Ok(())
}

fn create_combobox(hwnd: HWND) -> Result<()> {
    let hwnd = unsafe {
        CreateWindowExW(
            WS_EX_STATICEDGE,
            WC_COMBOBOXW,
            None,
            WINDOW_STYLE((CBS_DROPDOWNLIST | CBS_HASSTRINGS | CBS_SORT) as _)
                | WS_CHILD
                | WS_VISIBLE
                | WS_VSCROLL,
            340,
            12,
            227,
            200,
            hwnd,
            HMENU(ID_COMBO as _),
            None,
            None,
        )?
    };

    SpeechSynthesizer::AllVoices()?
        .into_iter()
        .try_for_each(|v| -> Result<()> {
            let name = v.DisplayName()?;
            unsafe { SendMessageW(hwnd, CB_ADDSTRING, None, LPARAM(name.as_ptr() as _)) };
            Ok(())
        })?;

    let default_voice = SpeechSynthesizer::DefaultVoice()?.DisplayName()?;
    unsafe {
        SendMessageW(
            hwnd,
            CB_SELECTSTRING,
            None,
            LPARAM(default_voice.as_ptr() as _),
        )
    };
    COMBOBOX_HWND.get_or_init(|| Hwnd::new(hwnd));
    Ok(())
}

fn create_edit(hwnd: HWND) -> Result<()> {
    let rc = unsafe {
        let mut rc = RECT::default();
        GetClientRect(hwnd, &mut rc)?;
        rc
    };
    let hwnd = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("EDIT"),
            None,
            WINDOW_STYLE((ES_MULTILINE | ES_WANTRETURN | /*ES_AUTOHSCROLL|*/ ES_AUTOVSCROLL) as _)
                | WS_CHILD
                | WS_VISIBLE
                | WS_BORDER
                | WS_TABSTOP
                //| WS_HSCROLL,
            | WS_VSCROLL,
            0,
            80,
            rc.right,
            rc.bottom - 80,
            hwnd,
            None,
            GetModuleHandleW(None)?,
            None,
        )?
    };
    EDIT_HWND.get_or_init(|| Hwnd::new(hwnd));
    Ok(())
}

fn create_trackbar(hwnd: HWND) -> Result<()> {
    let hwnd = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("msctls_trackbar32"),
            w!("Track Bar"),
            WS_CHILD | WS_VISIBLE | WINDOW_STYLE(TBS_TOOLTIPS | TBS_AUTOTICKS),
            145,
            50,
            400,
            30,
            hwnd,
            HMENU(ID_TRACKBAR as _),
            None,
            None,
        )
    }?;
    unsafe { SendMessageW(hwnd, TBM_SETRANGE, WPARAM(1), LPARAM(makelong(5, 25) as _)) };
    unsafe { SendMessageW(hwnd, TBM_SETPAGESIZE, None, LPARAM(5)) };
    unsafe { SendMessageW(hwnd, TBM_SETTICFREQ, WPARAM(5), LPARAM(0)) };
    unsafe { SendMessageW(hwnd, TBM_SETPOS, WPARAM(1), LPARAM(10)) };
    TRACKBAR_HWND.get_or_init(|| Hwnd::new(hwnd));
    Ok(())
}

/// トラックバーを生成するためにコモンコントロールを初期化する
fn init_common_control() -> Result<()> {
    let icc = INITCOMMONCONTROLSEX {
        dwSize: size_of::<INITCOMMONCONTROLSEX>() as _,
        dwICC: ICC_BAR_CLASSES,
    };
    unsafe { InitCommonControlsEx(&icc).ok()? };
    Ok(())
}

/// 各種 UI を生成する
fn create(hwnd: HWND) -> Result<()> {
    init_common_control()?;
    create_play_button(hwnd)?;
    create_clear_button(hwnd)?;
    create_save_button(hwnd)?;
    create_edit(hwnd)?;
    create_combobox(hwnd)?;
    create_trackbar(hwnd)?;
    Ok(())
}

/// ウィンドウプロシージャ
unsafe extern "system" fn wnd_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            create(hwnd).ok();
        }
        WM_COMMAND => {
            command(hwnd, wparam).ok();
        }
        WM_PAINT => {
            paint(hwnd).ok();
        }
        WM_DESTROY => PostQuitMessage(0),
        _ => return DefWindowProcW(hwnd, msg, wparam, lparam),
    }
    LRESULT::default()
}

/// エントリーポイント
fn main() -> Result<()> {
    let wnd_class = WNDCLASSW {
        lpfnWndProc: Some(wnd_proc),
        lpszClassName: CLASS_NAME,
        hbrBackground: unsafe { GetSysColorBrush(COLOR_MENUBAR) },
        ..Default::default()
    };

    unsafe { RegisterClassW(&wnd_class) };

    let hwnd = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            CLASS_NAME,
            w!("speech"),
            WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_MINIMIZEBOX,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            600,
            480,
            None,
            None,
            None,
            None,
        )?
    };

    unsafe { ShowWindow(hwnd, SW_SHOW).ok()? };
    unsafe { UpdateWindow(hwnd).ok()? };

    let mut msg = MSG::default();

    loop {
        if !unsafe { GetMessageW(&mut msg, None, 0, 0) }.as_bool() {
            break;
        }
        unsafe {
            _ = TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
    Ok(())
}

/// ヘルパー関数
#[inline]
fn makelong(a: u16, b: u16) -> i32 {
    ((a as u32) | ((b as u32) << 16)) as i32
}

/// ヘルパー関数
#[inline]
fn loword(dword: u32) -> u16 {
    ((dword << 16) >> 16) as _
}
