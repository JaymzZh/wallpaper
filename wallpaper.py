#!/usr/bin/env python
# encoding:utf-8

from __future__ import print_function
import ctypes
from ctypes import wintypes
import win32con
from urllib import request
import xml.etree.ElementTree as eTree
import os
from os import path
import socket
import sys
import random
import time
import logging
from threading import Thread

__author__ = 'zhangmm'

logger = logging.getLogger('wallpapper')
logger.setLevel(logging.DEBUG)
# create file handler which logs even debug messages
fh = logging.FileHandler('wallpapper.log')
fh.setLevel(logging.DEBUG)
# create console handler with a debug log level
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
# create formatter and add it to the handlers
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
ch.setFormatter(formatter)
fh.setFormatter(formatter)
# add the handlers to logger
logger.addHandler(ch)
logger.addHandler(fh)


def download_picture():
    basedir = path.join(path.abspath(path.dirname(__file__)), 'wallpapers')
    if not path.exists(basedir):
        os.mkdir(basedir)
    logger.info('download picture begin:')
    count = 0
    for i in range(8, -1, -1):
        xmlurl = 'http://az517271.vo.msecnd.net/TodayImageService.svc/HPImageArchive?mkt=zh-cn&idx=%d' % i
        # this url supports ipv6, but cn.bing.com doesn't
        try:
            xmlhandle = request.urlopen(xmlurl, timeout=5)
            xmlresponse = xmlhandle.read()
            root = eTree.fromstring(xmlresponse)
        except socket.timeout:
            logger.error('timeout downloading image information.')
            continue

        datestr = root[0].text
        imgpath = path.join(basedir, '%s.jpg' % datestr)
        if not path.exists(imgpath):
            imgurl = root[6].text
            try:
                imgdata = request.urlopen(imgurl, timeout=5).read()
                if len(imgdata) > 100 * 1024:  # if tunet not authorized
                    with open(imgpath, 'wb') as imgfile:
                        imgfile.write(imgdata)
                    count += 1
            except socket.timeout:
                logger.error('timeout downloading wallpapers.')
                continue
    logger.info('download picture end, total download %s' % (count))


def set_wallpaper(picpath):
    if sys.platform == 'win32':
        import win32api, win32con, win32gui
        k = win32api.RegOpenKey(win32con.HKEY_CURRENT_USER, 'Control Panel\Desktop', 0, win32con.KEY_ALL_ACCESS)
        curpath = win32api.RegQueryValueEx(k, 'Wallpaper')[0]
        if curpath != picpath:
            win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, picpath, 1 + 2)
        win32api.RegCloseKey(k)
    else:
        import commands
        curpath = commands.getstatusoutput('gsettings get org.gnome.desktop.background picture-uri')[1][1:-1]
        if curpath != picpath:
            commands.getstatusoutput(
                'DISPLAY=:0 gsettings set org.gnome.desktop.background picture-uri "%s"' % (picpath))


def get_random_image():
    basedir = path.join(path.abspath(path.dirname(__file__)), 'wallpapers')
    if not path.exists(basedir):
        os.mkdir(basedir)
    paths = os.walk(basedir)
    for root, dirs, files in paths:
        files = files
    validpath = random.choice(files)
    validpath = path.join(basedir, validpath)
    logger.info(validpath)
    return validpath


day_runed = '2015-10-20'
is_runed = False


def download_picture_perday():
    global day_runed, is_runed
    day_now = time.strftime('%Y-%m-%d', time.localtime(time.time()))
    if day_runed != day_now and not is_runed:
        thread = Thread(target=download_picture)
        thread.start()
        day_runed = day_now
        is_runed = True


def add_hotkey():
    # Register hotkeys
    byref = ctypes.byref
    user32 = ctypes.windll.user32

    HOTKEYS = {
        1: (win32con.VK_F4, win32con.MOD_WIN),
        2: (win32con.VK_RETURN, win32con.MOD_WIN)
    }

    def handle_win_home():
        set_wallpaper('force.jpg')

    def handle_win_f4():
        user32.PostQuitMessage(0)
        sys.exit()

    HOTKEY_ACTIONS = {
        1: handle_win_f4,
        2: handle_win_home
    }

    #
    # RegisterHotKey takes:
    #  Window handle for WM_HOTKEY messages (None = this thread)
    #  arbitrary id unique within the thread
    #  modifiers (MOD_SHIFT, MOD_ALT, MOD_CONTROL, MOD_WIN)
    #  VK code (either ord ('x') or one of win32con.VK_*)
    #
    for id, (vk, modifiers) in HOTKEYS.items():
        logger.debug("Registering id %s for key %s" % (id, vk))
        if not user32.RegisterHotKey(None, id, modifiers, vk):
            logger.info("Unable to register id %s" % id)

    #
    # Home-grown Windows message loop: does
    #  just enough to handle the WM_HOTKEY
    #  messages and pass everything else along.
    #
    try:
        msg = wintypes.MSG()
        while user32.GetMessageA(byref(msg), None, 0, 0) != 0:
            if msg.message == win32con.WM_HOTKEY:
                action_to_take = HOTKEY_ACTIONS.get(msg.wParam)
                if action_to_take:
                    action_to_take()

            user32.TranslateMessage(byref(msg))
            user32.DispatchMessageA(byref(msg))
    except Exception as e:
        logger.error(e.args)
    finally:
        for id in HOTKEYS.keys():
            user32.UnregisterHotKey(None, id)


if __name__ == '__main__':
    add_hotkey()

    while 1:
        download_picture_perday()
        picpath = get_random_image()
        set_wallpaper(picpath)
        t = random.randint(10, 30)
        t = t * 60
        time.sleep(t)
