using SharpDX;
using SharpDX.DirectSound;
using SharpDX.Multimedia;
using System;
using System.Collections.Generic;
using System.IO;

public class DXSoundPlayer : IDisposable
{
    private DirectSound _device;
    private Dictionary<string, SecondarySoundBuffer> _buffers = new Dictionary<string, SecondarySoundBuffer>();

    public DXSoundPlayer(IntPtr windowHandle)
    {
        _device = new DirectSound();
        _device.SetCooperativeLevel(windowHandle, CooperativeLevel.Priority);
    }

    /// <summary>
    /// Загружает WAV-файл в звуковой буфер с заданным именем.
    /// </summary>
    public void LoadWav(string name, string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("WAV file not found", path);

        // если был буфер с таким именем — удаляем
        if (_buffers.ContainsKey(name))
        {
            _buffers[name].Dispose();
            _buffers.Remove(name);
        }

        using (var soundStream = new SoundStream(File.OpenRead(path)))
        {
            WaveFormat format = soundStream.Format;

            using (var dataStream = soundStream.ToDataStream())
            {
                int size = (int)dataStream.Length;
                byte[] audioData = new byte[size];
                dataStream.Read(audioData, 0, size);

                var desc = new SoundBufferDescription
                {
                    BufferBytes = size,
                    Format = format,
                    Flags =
                        BufferFlags.ControlVolume |
                        BufferFlags.GlobalFocus |
                        BufferFlags.GetCurrentPosition2
                };

                var buffer = new SecondarySoundBuffer(_device, desc);

                buffer.Write(audioData, 0, LockFlags.EntireBuffer);

                _buffers[name] = buffer;
            }
        }
    }

    /// <summary>
    /// Проигрывает звук по имени.
    /// loop = true — бесконечный цикл.
    /// loop = false — один раз.
    /// </summary>
    public void Play(string name, bool loop)
    {
        if (!_buffers.ContainsKey(name))
            throw new ArgumentException($"Buffer '{name}' not loaded.");

        var buf = _buffers[name];

        buf.Stop();
        buf.CurrentPosition = 0;

        var flags = loop ? PlayFlags.Looping : PlayFlags.None;

        // ВАЖНО: нужно указать priority (обычно 0) и флаги
        buf.Play(0, flags);
    }

    /// <summary>
    /// Останавливает звук по имени.
    /// </summary>
    public void Stop(string name)
    {
        if (_buffers.ContainsKey(name))
            _buffers[name].Stop();
    }

    /// <summary>
    /// Останавливает ВСЕ звуки.
    /// </summary>
    public void StopAll()
    {
        foreach (var b in _buffers.Values)
            b.Stop();
    }

    public void Dispose()
    {
        StopAll();
        foreach (var b in _buffers.Values)
            b.Dispose();
        _buffers.Clear();

        _device?.Dispose();
    }
}