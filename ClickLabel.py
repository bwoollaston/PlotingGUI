from matplotlib.text import Text
import matplotlib.pyplot as plt
class LabelHandler:
    def __init__(self, ax):
        self.ax = ax
        self.labels = []
        self.current_label = None
        self.cid_press = ax.figure.canvas.mpl_connect('button_press_event', self.on_press)
        self.cid_release = ax.figure.canvas.mpl_connect('button_release_event', self.on_release)
        self.cid_motion = ax.figure.canvas.mpl_connect('motion_notify_event', self.on_motion)

    def on_press(self, event):
        if event.inaxes != self.ax:
            return
        self.current_label = plt.text(event.xdata, event.ydata, 'Click and Drag', color='red')
        self.labels.append(self.current_label)
        plt.draw()

    def on_motion(self, event):
        if self.current_label is None or event.inaxes != self.ax:
            return
        self.current_label.set_position((event.xdata, event.ydata))
        plt.draw()

    def on_release(self, event):
        self.current_label = None
        plt.draw()