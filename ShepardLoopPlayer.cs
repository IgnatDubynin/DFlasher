using System;
using System.IO;
using System.Windows.Forms;
using SharpDX;
using SharpDX.DirectSound;
using SharpDX.Multimedia;

namespace DFlasher
{
    public class ShepardLoopPlayer : IDisposable
    {
        private readonly DirectSound _device;
        private readonly SecondarySoundBuffer _bufferUp;
        private readonly SecondarySoundBuffer _bufferDown;

        private bool _isUpPlaying;
        private bool _isDownPlaying;

        public ShepardLoopPlayer(IntPtr windowHandle, string wavUpPath, string wavDownPath)
        {
            // Инициализация устройства DirectSound
            _device = new DirectSound();
            _device.SetCooperativeLevel(windowHandle, CooperativeLevel.Priority);

            _bufferUp = CreateLoopBufferFromWav(wavUpPath);
            _bufferDown = CreateLoopBufferFromWav(wavDownPath);
        }

        private SecondarySoundBuffer CreateLoopBufferFromWav(string path)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException("WAV file not found", path);

            using (var soundStream = new SoundStream(File.OpenRead(path)))
            {
                WaveFormat waveFormat = soundStream.Format;

                // читаем все данные WAV в память
                using (var dataStream = soundStream.ToDataStream())
                {
                    int bytes = (int)dataStream.Length;
                    byte[] audioData = new byte[bytes];
                    dataStream.Read(audioData, 0, bytes);

                    var bufferDesc = new SoundBufferDescription
                    {
                        BufferBytes = bytes,
                        Format = waveFormat,
                        Flags = BufferFlags.ControlVolume | BufferFlags.GlobalFocus
                    };

                    // ВАЖНО: используем КОНСТРУКТОР с 2 аргументами
                    var buffer = new SecondarySoundBuffer(_device, bufferDesc);

                    // заливаем данные в буфер целиком
                    buffer.Write(audioData, 0, LockFlags.EntireBuffer);

                    return buffer;
                }
            }
        }

        /// <summary>
        /// dir &lt; 0  -> играть down
        /// dir &gt; 0  -> играть up
        /// dir = 0  -> остановить всё
        /// </summary>
        public void PlayForDirection(int dir)
        {
            if (dir < 0)
            {
                // вниз: включаем down, выключаем up
                StopUpInternal();

                if (!_isDownPlaying)
                {
                    _bufferDown.CurrentPosition = 0;
                    _bufferDown.Play(0, PlayFlags.Looping);
                    _isDownPlaying = true;
                }
            }
            else if (dir > 0)
            {
                // вверх: включаем up, выключаем down
                StopDownInternal();

                if (!_isUpPlaying)
                {
                    _bufferUp.CurrentPosition = 0;
                    _bufferUp.Play(0, PlayFlags.Looping);
                    _isUpPlaying = true;
                }
            }
            else
            {
                // 0 – всё стоп
                StopUpInternal();
                StopDownInternal();
            }
        }

        private void StopUpInternal()
        {
            if (_isUpPlaying)
            {
                _bufferUp.Stop();
                _isUpPlaying = false;
            }
        }

        private void StopDownInternal()
        {
            if (_isDownPlaying)
            {
                _bufferDown.Stop();
                _isDownPlaying = false;
            }
        }

        public void Dispose()
        {
            StopUpInternal();
            StopDownInternal();

            if (_bufferUp != null) _bufferUp.Dispose();
            if (_bufferDown != null) _bufferDown.Dispose();
            if (_device != null) _device.Dispose();
        }
    }
}