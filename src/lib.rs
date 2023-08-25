//! # CeVIO
//! Rust から CeVIO/CeVIO AI の [COM コンポーネント API](https://cevio.jp/guide/cevio_ai/interface/) を使用するためのライブラリです
//!
//! ## 使い方
//!
//! ```no_run
//! use cevio::CeVIO;
//! let cevio = CeVIO::new().unwrap();
//!
//! cevio.start_host(false).unwrap(); // CeVIO AI を起動
//! cevio.set_cast("花隈千冬").unwrap(); // 【必須】キャストを設定
//! cevio.set_volume(100).unwrap(); // 音量を設定
//! cevio.set_tone(50).unwrap(); // 音の高さを設定
//! cevio.output_wave_to_file( // .wav 形式で出力
//!     "初めまして。花隈千冬です。よろしくお願いします。",
//!     r"E:\file.wav", // 出力パスを設定
//! ).unwrap();
//! ```
//!
//! 詳しくはこちら: [struct CeVIO](./struct.CeVIO.html)

use anyhow::Context as _;
use windows::Win32::System::Com::VARIANT;

mod com;
pub mod error;
mod initialize;
mod variant_ext;

use com::ComObject;
use initialize::Initialize;
use variant_ext::VariantExt;

pub struct CeVIO {
    _init: Initialize,
    talker: ComObject,
    controller: ComObject,
}

fn make_error_message(method_name: &str, fn_name: &str) -> String {
    format!("Failed to call `{method_name}` in fn `{fn_name}`")
}

impl CeVIO {
    /// CeVIO AI 用インスタンスを作成します。（`CeVIO::new_cevio_ai` と同じです。）
    ///
    /// CeVIO を使用する場合は `CeVIO::new_cevio()` を使用してください。
    pub fn new() -> error::Result<Self> {
        Ok(Self {
            _init: Initialize::new().map_err(error::CeVIOError)?,
            talker: ComObject::new("CeVIO.Talk.RemoteService2.Talker2")
                .map_err(|e| e.into())
                .map_err(error::CeVIOError)?,
            controller: ComObject::new("CeVIO.Talk.RemoteService2.ServiceControl2")
                .map_err(|e| e.into())
                .map_err(error::CeVIOError)?,
        })
    }

    /// CeVIO 用インスタンスを作成します。
    ///
    /// CeVIO AI を使用する場合は `CeVIO::new_cevio_ai()` を使用してください。
    pub fn new_cevio() -> error::Result<Self> {
        Ok(Self {
            _init: Initialize::new().map_err(error::CeVIOError)?,
            talker: ComObject::new("CeVIO.Talk.RemoteService.Talker")
                .map_err(|e| e.into())
                .map_err(error::CeVIOError)?,
            controller: ComObject::new("CeVIO.Talk.RemoteService.ServiceControl")
                .map_err(|e| e.into())
                .map_err(error::CeVIOError)?,
        })
    }

    /// CeVIO AI 用インスタンスを作成します。
    ///
    /// CeVIO を使用する場合は `CeVIO::new_cevio()` を使用してください。
    pub fn new_cevio_ai() -> error::Result<Self> {
        Ok(Self {
            _init: Initialize::new().map_err(error::CeVIOError)?,
            talker: ComObject::new("CeVIO.Talk.RemoteService2.Talker2")
                .map_err(|e| e.into())
                .map_err(error::CeVIOError)?,
            controller: ComObject::new("CeVIO.Talk.RemoteService2.ServiceControl2")
                .map_err(|e| e.into())
                .map_err(error::CeVIOError)?,
        })
    }

    /// 【CeVIO Creative Studio】を起動します。起動済みなら何もしません。
    ///
    /// 引数：
    ///
    /// 　noWait - trueは起動のみ行います。アクセス可能かどうかはIsHostStartedで確認します。
    ///
    /// 　　　　　　falseは起動後に外部からアクセス可能になるまで制御を戻しません。
    ///
    /// 戻り値：
    ///
    /// 　 0：成功。起動済みの場合も含みます。
    ///
    /// 　-1：インストール状態が不明。
    ///
    /// 　-2：実行ファイルが見つからない。
    ///
    /// 　-3：プロセスの起動に失敗。
    ///
    /// 　-4：アプリケーション起動後、エラーにより終了。
    pub fn start_host(&self, no_wait: bool) -> error::Result<i32> {
        self.controller
            .invoke_method("StartHost", vec![VARIANT::from_bool(no_wait)])
            .with_context(|| make_error_message("invoke_method", "start_host"))
            .map_err(error::CeVIOError)?
            .to_i32()
            .with_context(|| make_error_message("to_i32", "start_host"))
            .map_err(error::CeVIOError)
    }

    /// 【CeVIO Creative Studio】に終了を要求します。
    ///
    /// 引数：
    ///
    /// 　mode - 処理モード。
    ///
    /// 　 0：【CeVIO AI】が編集中の場合、保存や終了キャンセルが可能。
    pub fn close_host(&self, mode: i32) -> error::Result<()> {
        self.controller
            .invoke_method("CloseHost", vec![VARIANT::from_i32(mode)])
            .with_context(|| make_error_message("invoke_method", "close_host"))
            .map_err(error::CeVIOError)?;
        Ok(())
    }

    /// 【CeVIO Creative Studio】のバージョンを取得します。
    pub fn get_host_version(&self) -> error::Result<String> {
        self.controller
            .get_property("HostVersion", None)
            .with_context(|| make_error_message("get_property", "get_host_version"))
            .map_err(error::CeVIOError)?
            .to_string()
            .with_context(|| make_error_message("to_string", "get_host_version"))
            .map_err(error::CeVIOError)
    }

    /// このライブラリのバージョンを取得します。
    pub fn get_interface_version(&self) -> error::Result<String> {
        self.controller
            .get_property("InterfaceVersion", None)
            .with_context(|| make_error_message("get_property", "get_interface_version"))
            .map_err(error::CeVIOError)?
            .to_string()
            .with_context(|| make_error_message("to_string", "get_interface_version"))
            .map_err(error::CeVIOError)
    }

    /// 【CeVIO Creative Studio】にアクセス可能かどうか取得します。
    pub fn get_is_host_started(&self) -> error::Result<bool> {
        self.controller
            .get_property("InterfaceVersion", None)
            .with_context(|| make_error_message("get_property", "get_is_host_started"))
            .map_err(error::CeVIOError)?
            .to_bool()
            .with_context(|| make_error_message("to_bool", "get_is_host_started"))
            .map_err(error::CeVIOError)
    }

    /// 音の大きさ（0～100）を取得します。
    pub fn get_volume(&self) -> error::Result<i32> {
        self.talker
            .get_property("Volume", None)
            .with_context(|| make_error_message("get_property", "get_volume"))
            .map_err(error::CeVIOError)?
            .to_i32()
            .with_context(|| make_error_message("to_i32", "get_volume"))
            .map_err(error::CeVIOError)
    }

    /// 音の大きさ（0～100）を設定します。
    pub fn set_volume(&self, volume: i32) -> error::Result<()> {
        self.talker
            .set_property("Volume", None, VARIANT::from_i32(volume))
            .with_context(|| make_error_message("set_property", "set_volume"))
            .map_err(error::CeVIOError)
    }

    /// 話す速さ（0～100）を取得します。
    pub fn get_speed(&self) -> error::Result<i32> {
        self.talker
            .get_property("Speed", None)
            .with_context(|| make_error_message("get_property", "get_speed"))
            .map_err(error::CeVIOError)?
            .to_i32()
            .with_context(|| make_error_message("to_i32", "get_speed"))
            .map_err(error::CeVIOError)
    }

    /// 話す速さ（0～100）を設定します。
    pub fn set_speed(&self, speed: i32) -> error::Result<()> {
        self.talker
            .set_property("Speed", None, VARIANT::from_i32(speed))
            .with_context(|| make_error_message("set_property", "set_speed"))
            .map_err(error::CeVIOError)
    }

    /// 音の高さ（0～100）を取得します。
    pub fn get_tone(&self) -> error::Result<i32> {
        self.talker
            .get_property("Tone", None)
            .with_context(|| make_error_message("get_property", "get_tone"))
            .map_err(error::CeVIOError)?
            .to_i32()
            .with_context(|| make_error_message("to_i32", "get_tone"))
            .map_err(error::CeVIOError)
    }

    /// 音の高さ（0～100）を設定します。
    pub fn set_tone(&self, tone: i32) -> error::Result<()> {
        self.talker
            .set_property("Tone", None, VARIANT::from_i32(tone))
            .with_context(|| make_error_message("set_property", "set_tone"))
            .map_err(error::CeVIOError)
    }

    /// 抑揚（0～100）を取得します。
    pub fn get_tone_scale(&self) -> error::Result<i32> {
        self.talker
            .get_property("ToneScale", None)
            .with_context(|| make_error_message("get_property", "get_tone_scale"))
            .map_err(error::CeVIOError)?
            .to_i32()
            .with_context(|| make_error_message("to_i32", "get_tone_scale"))
            .map_err(error::CeVIOError)
    }

    /// 抑揚（0～100）を設定します。
    pub fn set_tone_scale(&self, tone_scale: i32) -> error::Result<()> {
        self.talker
            .set_property("ToneScale", None, VARIANT::from_i32(tone_scale))
            .with_context(|| make_error_message("set_property", "set_tone_scale"))
            .map_err(error::CeVIOError)
    }

    /// 声質（0～100）を取得します。
    pub fn get_alpha(&self) -> error::Result<i32> {
        self.talker
            .get_property("Alpha", None)
            .with_context(|| make_error_message("get_property", "get_alpha"))
            .map_err(error::CeVIOError)?
            .to_i32()
            .with_context(|| make_error_message("to_i32", "get_alpha"))
            .map_err(error::CeVIOError)
    }

    /// 声質（0～100）を設定します。
    pub fn set_alpha(&self, alpha: i32) -> error::Result<()> {
        self.talker
            .set_property("Alpha", None, VARIANT::from_i32(alpha))
            .with_context(|| make_error_message("set_property", "set_alpha"))
            .map_err(error::CeVIOError)
    }

    /// キャストを取得します。
    pub fn get_cast(&self) -> error::Result<String> {
        self.talker
            .get_property("Cast", None)
            .with_context(|| make_error_message("get_property", "get_cast"))
            .map_err(error::CeVIOError)?
            .to_string()
            .with_context(|| make_error_message("to_string", "get_cast"))
            .map_err(error::CeVIOError)
    }

    /// キャストを設定します。
    pub fn set_cast(&self, cast: &str) -> error::Result<()> {
        self.talker
            .set_property("Cast", None, VARIANT::from_str(cast))
            .with_context(|| make_error_message("set_property", "set_cast"))
            .map_err(error::CeVIOError)
    }

    /// 利用可能なキャスト名を取得します。
    ///
    /// 備考：
    ///
    /// 　キャストの取り揃えは、インストールされている音源によります。
    ///
    /// 注意点：
    ///
    /// 　型は、Visual C++環境でスマートポインタを利用する場合、下記に置き換えられます。
    ///
    /// 　IStringArray2Ptr
    pub fn get_available_casts(&self) -> error::Result<String> {
        self.talker
            .get_property("AvailableCasts", None)
            .with_context(|| make_error_message("get_property", "get_available_casts"))
            .map_err(error::CeVIOError)?
            .to_string()
            .with_context(|| make_error_message("to_string", "get_available_casts"))
            .map_err(error::CeVIOError)
    }

    /// 指定したセリフの再生を開始します。
    ///
    /// 引数：
    ///
    /// 　text - セリフ。
    ///
    /// 戻り値：
    ///
    /// 　再生状態を表すオブジェクト。
    ///
    /// 備考：
    ///
    /// 　再生終了を待たずに処理が戻ります。
    ///
    /// 　再生終了を待つには戻り値（ISpeakingState2）のWaitを呼び出します。
    ///
    /// 注意点：
    ///
    /// 　型は、Visual C++環境でスマートポインタを利用する場合、下記に置き換えられます。
    ///
    /// 　ISpeakingState2Ptr
    pub fn speak(&self, text: &str) -> error::Result<()> {
        self.talker
            .invoke_method("Speak", vec![VARIANT::from_str(text)])
            .with_context(|| make_error_message("invoke_method", "speak"))
            .map_err(error::CeVIOError)?;
        Ok(())
    }

    /// 指定したセリフの音素単位のデータを取得します。
    ///
    /// 引数：
    ///
    /// 　text - セリフ。
    ///
    /// 戻り値：
    ///
    /// 　音素単位のデータ。
    ///
    /// 備考：
    ///
    /// 　リップシンク等に利用できます。
    ///
    /// 注意点：
    ///
    /// 　型は、Visual C++環境でスマートポインタを利用する場合、下記に置き換えられます。
    ///
    /// 　IPhonemeDataArray2Ptr
    pub fn get_phonemes(&self, text: &str) -> error::Result<()> {
        self.talker
            .invoke_method("GetPhonemes", vec![VARIANT::from_str(text)])
            .with_context(|| make_error_message("invoke_method", "get_phonemes"))
            .map_err(error::CeVIOError)?;
        Ok(())
    }

    /// 指定したセリフをWAVファイルとして出力します。
    ///
    /// 引数：
    ///
    /// 　text - セリフ。
    ///
    /// 　path - 出力先パス。
    ///
    /// 戻り値：
    ///
    /// 　成功した場合はtrue。それ以外の場合はfalse。
    ///
    /// 備考：
    ///
    /// 　出力形式はサンプリングレート48kHz, ビットレート16bit, モノラルです。
    pub fn output_wave_to_file(&self, text: &str, path: &str) -> error::Result<()> {
        self.talker
            .invoke_method(
                "OutputWaveToFile",
                vec![VARIANT::from_str(text), VARIANT::from_str(path)],
            )
            .with_context(|| make_error_message("invoke_method", "speak"))
            .map_err(error::CeVIOError)?;
        Ok(())
    }
}
