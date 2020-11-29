#!/usr/bin/python3.6
# tested on python 3.6

import pdftotext as ptt
import argparse
from urllib.parse import quote_plus
from requests import get
from requests.exceptions import ConnectionError
from os import path
import re


class PlaysoundException(Exception):
        pass


class PlaySound:
    def settings(self, sound, block = True):
        self.sound = sound
        self.block = block
        
    def play(self):
        from platform import system
        system = system()
        if system == 'Windows':
            self._playsoundWin()
        elif system == 'Darwin':
            self._playsoundOSX()
        else:
            self._playsound_nix()
        del system

    def _playsound_win(self):
        """
        Utilizes windll.winmm. Tested and known to work with MP3 and WAVE on
        Windows 7 with Python 2.7. Probably works with more file formats.
        Probably works on Windows XP thru Windows 10. Probably works with all
        versions of Python.

        Inspired by (but not copied from) Michael Gundlach <gundlach@gmail.com>'s mp3play:
        https://github.com/michaelgundlach/mp3play
        I never would have tried using windll.winmm without seeing his code.
        """
        from ctypes import c_buffer, windll
        from random import random
        from time import sleep
        from sys import getfilesystemencoding

        def win_command(*command):
            buf = c_buffer(255)
            command = ' '.join(command).encode(getfilesystemencoding())
            errorCode = int(windll.winmm.mciSendStringA(command, buf, 254, 0))
            if errorCode:
                errorBuffer = c_buffer(255)
                windll.winmm.mciGetErrorStringA(errorCode, errorBuffer, 254)
                exceptionMessage = ('\n    Error ' + str(errorCode) + ' for command:'
                                    '\n        ' + command.decode() +
                                    '\n    ' + errorBuffer.value.decode())
                raise PlaysoundException(exceptionMessage)
            return buf.value

        alias = 'playsound_' + str(random())
        winCommand('open "' + self.sound + '" alias', alias)
        winCommand('set', alias, 'time format milliseconds')
        durationInMS = winCommand('status', alias, 'length')
        winCommand('play', alias, 'from 0 to', durationInMS.decode())

        if block:
            sleep(float(durationInMS) / 1000.0)

    def _playsound_osx(self):
        '''
        Utilizes AppKit.NSSound. Tested and known to work with MP3 and WAVE on
        OS X 10.11 with Python 2.7. Probably works with anything QuickTime supports.
        Probably works on OS X 10.5 and newer. Probably works with all versions of
        Python.

        Inspired by (but not copied from) Aaron's Stack Overflow answer here:
        http://stackoverflow.com/a/34568298/901641

        I never would have tried using AppKit.NSSound without seeing his code.
        '''
        from AppKit import NSSound
        from Foundation import NSURL
        from time import sleep

        if '://' not in self.sound:
            if not self.sound.startswith('/'):
                from os import getcwd
                self.sound = getcwd() + '/' + self.sound
            self.sound = 'file://' + self.sound
        url = NSURL.URLWithString_(self.sound)
        nssound = NSSound.alloc().initWithContentsOfURL_byReference_(url, True)
        if not nssound:
            raise IOError('Unable to load sound named: ' + self.sound)
        nssound.play()

        if self.block:
            sleep(nssound.duration())

    def _playsound_nix(self):
        """Play a sound using GStreamer.
        Inspired by this:
        https://gstreamer.freedesktop.org/documentation/tutorials/playback/playbin-usage.html
        """
        if not self.block:
            raise NotImplementedError(
                "block=False cannot be used on this platform yet")

        # pathname2url escapes non-URL-safe characters
        import os
        try:
            from urllib.request import pathname2url
        except ImportError:
            # python 2
            from urllib import pathname2url

        import gi
        gi.require_version('Gst', '1.0')
        from gi.repository import Gst

        Gst.init(None)

        playbin = Gst.ElementFactory.make('playbin', 'playbin')
        if self.sound.startswith(('http://', 'https://')):
            playbin.props.uri = self.sound
        else:
            playbin.props.uri = 'file://' + pathname2url(os.path.abspath(self.sound))
            
        self.is_playing = True
        set_result = playbin.set_state(Gst.State.PLAYING)
        if set_result != Gst.StateChangeReturn.ASYNC:
            self.is_playing = False
            raise PlaysoundException(
                "playbin.set_state returned " + repr(set_result))

        # FIXME: use some other bus method than poll() with block=False
        # https://lazka.github.io/pgi-docs/#Gst-1.0/classes/Bus.html
        bus = playbin.get_bus()
        bus.poll(Gst.MessageType.EOS, Gst.CLOCK_TIME_NONE)
        playbin.set_state(Gst.State.NULL)
        self.is_playing = False


class CmdLineArgs:
    def __init__(self):
        self.parse = argparse.ArgumentParser()
        self.parse.add_argument("-f", "--file", help="PDF file to convert to audio", type=str)
        self.parse.add_argument("-o", "--output", help="file to save to", type=str)
        self.args = self.parse.parse_args()
        self.check_file()
        self.pdf_file = self.args.file
        self.process_file()
            
    def check_file(self):
        self.safe = False
        if self.args.file is None:
            exit("-f/--file argument is required")
        elif not self.args.file.endswith("pdf"):
            exit("Please provide a file with extension .pdf")
        else:
            self.safe = True
            
    def is_file_safe(self):
        return self.safe
        
    def process_file(self):
        with open(self.pdf_file, "rb") as f:
            self.pdf = ptt.PDF(f)
            
    def get_page(self, page):
        if page <= self.num_pages():
            if page > 0:
                page = page-1
            if page < 0:
                page = 0
            return self.pdf[page]
        else:
            exit("Page %d not found. Available pages is %d" % (page, self.num_pages()))
        
    def num_pages(self):
        return len(self.pdf)
        
    def __len__(self):
        return self.num_pages()


class TextToAudio(CmdLineArgs):
    def __init__(self):
        super().__init__()
        self.play_sound = PlaySound()
    
    def word_cleaner(self, sentence):
        return re.sub("<.+>", "{html}", sentence.strip())
    
    def audio(self, text):
        try:
            req = get("http://localhost:59125/process?INPUT_TEXT=%s&INPUT_TYPE=TEXT&OUTPUT_TYPE=AUDIO&LOCALE=en_US&AUDIO=WAVE_FILE" % 
                (quote_plus(text)))
            self.outfile = "/tmp/pdftoaudio.wav" if self.args.output is None else path.abspath(self.args.output)
            if not self.outfile.endswith("wav"):
                self.outfile = self.outfile + ".wav"
            outputfile = open(self.outfile, "wb")
            outputfile.write(req.content)
            outputfile.close()
        except ConnectionError as err:
            exit("**Error downloading file**\n %s" % re.search("'<.+'", str(err)).group(0))
            
    def play(self):
        self.play_sound.settings(self.outfile)
        self.play_sound.play()
        
    def start(self):
        page_num = 1
        line_num = 1
        for page in self.pdf:
            print("Processing page", page_num)
            page_num = page_num + 1
            lines = self.word_cleaner(page).split("\n")
            for line in lines:
                print("Line", line_num)
                self.audio(self.word_cleaner(line))
                self.play()
                line_num = line_num + 1
            line_num = 1
        

class PdfToAudio:
    def __init__(self):
        tta = TextToAudio()
        #   tta.audio("Hello world")
        tta.start()


if __name__ == "__main__":
    PdfToAudio()
