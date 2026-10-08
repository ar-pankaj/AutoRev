# -*- coding: utf-8 -*-
"""Splash screen shown while FilterMan loads."""
from __future__ import unicode_literals
import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

from System.Windows import (
    Window, WindowStyle, WindowStartupLocation, ResizeMode,
    HorizontalAlignment, VerticalAlignment, FontWeights
)
from System.Windows.Controls import Image, Grid, TextBlock
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
from System.Windows.Media import (
    Color, SolidColorBrush, RenderOptions, BitmapScalingMode,
    FontFamily, ColorConverter
)
from System.Windows.Media import Stretch
from System import Uri, UriKind
from System.Windows.Threading import Dispatcher, DispatcherPriority

_W = 440
_H = 280

_BG_LIGHT = Color.FromArgb(255, 240, 240, 240)
_BG_DARK  = Color.FromArgb(255,  28,  28,  28)


class SplashWindow(Window):
    def __init__(self, image_path, dark_mode=False):
        self.WindowStyle           = WindowStyle.None
        self.AllowsTransparency    = False
        self.ResizeMode            = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Width                 = _W
        self.Height                = _H
        self.Topmost               = True

        bg = _BG_DARK if dark_mode else _BG_LIGHT
        self.Background = SolidColorBrush(bg)

        # Load synchronously so PixelWidth/PixelHeight are available immediately
        bmp = BitmapImage()
        bmp.BeginInit()
        bmp.UriSource   = Uri(image_path, UriKind.Absolute)
        bmp.CacheOption = BitmapCacheOption.OnLoad
        bmp.EndInit()
        try:
            bmp.Freeze()
        except Exception:
            pass

        # Fit window exactly to the image's natural pixel size
        w = bmp.PixelWidth  if bmp.PixelWidth  > 0 else _W
        h = bmp.PixelHeight if bmp.PixelHeight > 0 else _H
        self.Width  = float(w)
        self.Height = float(h)

        img         = Image()
        img.Source  = bmp
        img.Stretch = Stretch.Fill  # no letterboxing - image fills the window exactly
        RenderOptions.SetBitmapScalingMode(img, BitmapScalingMode.HighQuality)
        grid        = Grid()
        grid.Children.Add(img)
        self.Content = grid


def show_splash(image_path, dark_mode=False):
    """Display splash window and return it. Caller is responsible for closing."""
    splash = SplashWindow(image_path, dark_mode=dark_mode)
    splash.Show()
    Dispatcher.CurrentDispatcher.Invoke(
        lambda: None,
        DispatcherPriority.Background
    )
    return splash


def close_after(splash_win, delay_ms):
    """Close splash_win after delay_ms on the UI thread. Returns timer (keep reference)."""
    from System.Windows.Threading import DispatcherTimer
    from System import TimeSpan

    timer = DispatcherTimer()
    timer.Interval = TimeSpan.FromMilliseconds(delay_ms)

    def _tick(s, a):
        timer.Stop()
        try:
            splash_win.Close()
        except Exception:
            pass

    timer.Tick += _tick
    timer.Start()
    return timer
