using Microsoft.CognitiveServices.Speech;
using Microsoft.CognitiveServices.Speech.Audio;
using System;
using System.Media;
using System.Threading.Tasks;
using System.Windows;

namespace MicrosoftAzureSpeechToText
{
	/// <summary>
	/// MainWindow.xamlのインタラクションロジック
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();
		}

		private void AppendLineLog(string log)
		{
			// 非同期処理から書き込むので Dispatcher.Invoke を使用
			Dispatcher.Invoke(()=>ResultTextBox.AppendText(log + Environment.NewLine));
		}

		private async void ExecuteButton_Click(object sender, RoutedEventArgs e)
		{
			// 入力内容をテキストボックスから取得
			var key = KeyTextBox.Text;
			var region = RegionTextBox.Text;
			var lang = LanguageTextBox.Text;
			var wavFilePath = WavFilePathTextBox.Text;

			try
			{
				// 音声ファイルが指定されているか確認するため再生する
				var wavPlayer = new SoundPlayer(wavFilePath);
				wavPlayer.Play();

				var stopRecognition = new TaskCompletionSource<int>();

				// 音声サービスを構成する
				var speechConfig = SpeechConfig.FromSubscription(key, region);
				AppendLineLog($"{speechConfig.Region} で音声サービスを使用する準備ができました。");

				// 音声認識言語の指定
				// 使用できる値一覧：https://docs.microsoft.com/ja-jp/azure/cognitive-services/speech-service/language-support?tabs=speechtotext#speech-to-text
				speechConfig.SpeechRecognitionLanguage = lang;

				// 入力を WAV ファイルとして設定
				using var audioConfig = AudioConfig.FromWavFileInput(wavFilePath);
				using var speechRecognizer = new SpeechRecognizer(speechConfig, audioConfig);

				// 解析結果が受信されたことを通知します。
				// このイベントは抽出が完了したものから随時発生します。
				speechRecognizer.Recognized += (s, e) =>
				{
					if (e.Result.Reason == ResultReason.RecognizedSpeech)
					{
						// 音声結果に認識されたテキストが含まれていることを示します。
						var time = TimeSpan.FromSeconds(e.Result.OffsetInTicks / 10000000).ToString(@"hh\:mm\:ss");
						var text = $"{time} {e.Result.Text}";
						AppendLineLog(text);
					}
					else if (e.Result.Reason == ResultReason.NoMatch)
					{
						// 音声を認識できなかったことを示します。
						AppendLineLog("音声を認識できませんでした。");
					}
				};

				// 音声認識が中断されたことを通知します。
				speechRecognizer.Canceled += (s, e) =>
				{
					AppendLineLog($"処理が終了しました。(Reason={e.Reason})");

					if (e.Reason == CancellationReason.Error)
					{
						AppendLineLog($"ErrorCode={e.ErrorCode}\r\n");
						AppendLineLog($"ErrorDetails={e.ErrorDetails}\r\n");
					}

					stopRecognition.TrySetResult(0);
				};

				// 継続的な処理を開始します。StopContinuousRecognitionAsync を使用して処理を停止します。
				await speechRecognizer.StartContinuousRecognitionAsync().ConfigureAwait(false);

				// 完了するのを待ちます。Task.WaitAny を使用して、タスクをルート化してください。
				Task.WaitAny(new[] { stopRecognition.Task });

				// 処理を停止します。
				await speechRecognizer.StopContinuousRecognitionAsync().ConfigureAwait(false);
			}
			catch (Exception ex)
			{
				// 何らかの例外が発生した場合はエラー内容を出力
				AppendLineLog(ex.Message);
			}

			MessageBox.Show("処理が終了しました。");
		}
	}
}
