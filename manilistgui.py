# -*- coding: utf-8 -*-

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
import manilist


class Root(BoxLayout):
    pass

class MyApp(App):
    title = 'My Application'

if __name__ == '__main__':
    MyApp().run()
    