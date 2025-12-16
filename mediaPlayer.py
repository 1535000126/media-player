import sys
import csv
import shutil
from pathlib import Path
from datetime import datetime

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QLabel, QPushButton, QFileDialog,
    QComboBox, QListWidget,
    QTimeEdit, QVBoxLayout, QHBoxLayout,
    QMessageBox, QCheckBox, QSplitter, QSlider
)
from PySide6.QtMultimedia import QMediaPlayer, QAudioOutput
from PySide6.QtMultimediaWidgets import QVideoWidget
from PySide6.QtCore import QUrl, QTimer, QTime, Qt
from PySide6.QtGui import QKeySequence, QShortcut


CSV_FILE = "time_ranges.csv"
SEGMENTS_DIR = Path("segments")
ORDER_FILE = "order.txt"  # per segment folder


# =====================
# 时间工具
# =====================

def hhmm_to_min(s: str) -> int:
    h, m = map(int, s.split(":"))
    return h * 60 + m

def min_to_hhmm(m: int) -> str:
    return f"{m // 60:02d}:{m % 60:02d}"

def qtime_to_min(t: QTime) -> int:
    return t.hour() * 60 + t.minute()

def segment_folder(s: int, e: int) -> str:
    return f"{min_to_hhmm(s).replace(':','_')}-{min_to_hhmm(e).replace(':','_')}"

def has_overlap(ranges, new_range):
    ns, ne = new_range
    return any(max(s, ns) < min(e, ne) for s, e in ranges)

def now_minutes() -> int:
    n = datetime.now()
    return n.hour * 60 + n.minute


# =====================
# 点播窗口（独立播放器 + 结束点播 + 真全屏）
# =====================

class OnDemandWindow(QMainWindow):
    """
    on_finish(terminated: bool)
      - terminated=False: 播完自然结束
      - terminated=True : 用户点击“结束点播”中断结束
    """
    def __init__(self, on_finish, get_volume_0_100, get_muted):
        super().__init__()
        self.setWindowTitle("点播")
        self.resize(1000, 700)

        self.on_finish = on_finish
        self.get_volume_0_100 = get_volume_0_100
        self.get_muted = get_muted

        self.playlist: list[Path] = []
        self.index = 0

        self.player = QMediaPlayer(self)
        self.audio = QAudioOutput(self)
        self.player.setAudioOutput(self.audio)

        self.video = QVideoWidget(self)
        self.player.setVideoOutput(self.video)
        self.player.mediaStatusChanged.connect(self.on_status)

        self.apply_volume_state()

        # 右侧控制区
        self.list_view = QListWidget()
        self.list_view.setDragDropMode(QListWidget.InternalMove)
        self.list_view.setDefaultDropAction(Qt.MoveAction)

        add_btn = QPushButton("添加视频")
        del_btn = QPushButton("删除选中")
        start_btn = QPushButton("开始点播")
        stop_btn = QPushButton("结束点播")
        fullscreen_btn = QPushButton("全屏")

        add_btn.clicked.connect(self.add_video)
        del_btn.clicked.connect(self.delete_selected)
        start_btn.clicked.connect(self.start_play)
        stop_btn.clicked.connect(self.stop_on_demand)
        fullscreen_btn.clicked.connect(self.enter_fullscreen)

        btn_row1 = QHBoxLayout()
        btn_row1.addWidget(add_btn)
        btn_row1.addWidget(del_btn)

        btn_row2 = QHBoxLayout()
        btn_row2.addWidget(start_btn)
        btn_row2.addWidget(stop_btn)
        btn_row2.addWidget(fullscreen_btn)

        right = QWidget()
        right_layout = QVBoxLayout()
        right_layout.addWidget(QLabel("点播列表（可拖动排序）"))
        right_layout.addWidget(self.list_view, 1)
        right_layout.addLayout(btn_row1)
        right_layout.addLayout(btn_row2)
        right.setLayout(right_layout)

        self.splitter = QSplitter(Qt.Horizontal)
        self.right_panel = right
        self.splitter.addWidget(self.video)
        self.splitter.addWidget(self.right_panel)
        self.splitter.setStretchFactor(0, 1)
        self.splitter.setStretchFactor(1, 0)

        self.setCentralWidget(self.splitter)

    def apply_volume_state(self):
        muted = bool(self.get_muted())
        vol = int(self.get_volume_0_100())
        self.audio.setMuted(muted)
        if not muted:
            self.audio.setVolume(max(0.0, min(1.0, vol / 100.0)))

    def set_volume_from_main(self, vol_0_100: int, muted: bool):
        self.audio.setMuted(bool(muted))
        if not muted:
            self.audio.setVolume(max(0.0, min(1.0, int(vol_0_100) / 100.0)))

    def toggle_pause(self):
        st = self.player.playbackState()
        if st == QMediaPlayer.PlayingState:
            self.player.pause()
        else:
            # Paused/Stopped 都当作继续播放（Stopped 时从当前 source 重新 play）
            self.player.play()

    def add_video(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "选择视频", "", "Video (*.mp4 *.mkv *.avi)")
        for p in paths:
            fp = Path(p)
            self.playlist.append(fp)
            self.list_view.addItem(fp.name)

    def delete_selected(self):
        r = self.list_view.currentRow()
        if r < 0:
            return
        self.list_view.takeItem(r)
        del self.playlist[r]
        if self.index >= len(self.playlist):
            self.index = max(0, len(self.playlist) - 1)

    def start_play(self):
        if self.list_view.count() == 0:
            return

        order = [self.list_view.item(i).text() for i in range(self.list_view.count())]
        name_to_path = {p.name: p for p in self.playlist}
        self.playlist = [name_to_path[n] for n in order if n in name_to_path]
        if not self.playlist:
            return

        self.index = 0
        self.enter_fullscreen()
        self.play_current()

    def play_current(self):
        self.apply_volume_state()
        self.player.setSource(QUrl.fromLocalFile(str(self.playlist[self.index])))
        self.player.play()

    def on_status(self, status):
        if status != QMediaPlayer.EndOfMedia:
            return
        self.index += 1
        if self.index >= len(self.playlist):
            self.close()
            self.on_finish(False)
        else:
            self.play_current()

    def stop_on_demand(self):
        self.player.stop()
        self.close()
        self.on_finish(True)

    def enter_fullscreen(self):
        self.showFullScreen()
        self.right_panel.hide()

    def exit_fullscreen(self):
        self.showNormal()
        self.right_panel.show()

    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Space:
            self.toggle_pause()
            return
        if e.key() == Qt.Key_Escape and self.isFullScreen():
            self.exit_fullscreen()
            return
        super().keyPressEvent(e)


# =====================
# 主窗口
# =====================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Time Scheduler Player")
        self.resize(1200, 750)

        SEGMENTS_DIR.mkdir(exist_ok=True)

        # ===== 音量状态（主控，0~100）=====
        self.volume = 50
        self.muted = False

        # ===== 调度/播放状态 =====
        self.time_ranges: list[tuple[int, int]] = []
        self.running = False
        self.current_segment: tuple[int, int] | None = None
        self.current_playlist: list[Path] = []
        self.current_index = 0
        self.waiting_next_segment: tuple[int, int] | None = None

        # 点播中断状态
        self.in_interrupt = False
        # (segment, playlist, index, position_ms, waiting_next_segment)
        self.interrupt_state = None
        self.ondemand_window: OnDemandWindow | None = None

        # ===== 播放器（主）=====
        self.player = QMediaPlayer(self)
        self.audio = QAudioOutput(self)
        self.player.setAudioOutput(self.audio)

        self.video = QVideoWidget(self)
        self.player.setVideoOutput(self.video)

        self.apply_volume_to_main_audio()
        self.player.mediaStatusChanged.connect(self.on_media_status)

        # ===== 定时器：真实时间检测 =====
        self.clock_timer = QTimer(self)
        self.clock_timer.timeout.connect(self.check_real_time)

        # ===== UI：时间段输入 =====
        self.start_t = QTimeEdit(QTime(0, 0))
        self.start_t.setDisplayFormat("HH:mm")
        self.end_t = QTimeEdit(QTime(0, 0))
        self.end_t.setDisplayFormat("HH:mm")

        # ===== UI：时间段选择 + 播放列表 =====
        self.segment_box = QComboBox()
        self.segment_box.currentIndexChanged.connect(self.refresh_playlist_view)

        self.playlist_view = QListWidget()
        self.playlist_view.setDragDropMode(QListWidget.InternalMove)
        self.playlist_view.setDefaultDropAction(Qt.MoveAction)
        self.playlist_view.model().rowsMoved.connect(
            lambda *args: self.save_ui_order_for_selected_segment()
        )

        # ===== UI：切换策略 =====
        self.graceful_switch_box = QCheckBox("播放完整视频后再切换")
        self.graceful_switch_box.setChecked(True)

        # ===== UI：音量 =====
        self.volume_slider = QSlider(Qt.Horizontal)
        self.volume_slider.setRange(0, 100)
        self.volume_slider.setValue(self.volume)
        self.volume_slider.valueChanged.connect(self.set_volume)

        mute_btn = QPushButton("静音/取消静音")
        mute_btn.clicked.connect(self.toggle_mute)

        # ===== UI：按钮 =====
        add_seg_btn = QPushButton("新增时间段")
        del_seg_btn = QPushButton("删除时间段")
        add_vid_btn = QPushButton("新增视频")
        del_vid_btn = QPushButton("删除视频")

        self.run_btn = QPushButton("开始运行")
        fullscreen_btn = QPushButton("全屏")
        self.ondemand_btn = QPushButton("点播")

        add_seg_btn.clicked.connect(self.add_segment)
        del_seg_btn.clicked.connect(self.delete_segment)
        add_vid_btn.clicked.connect(self.add_video)
        del_vid_btn.clicked.connect(self.delete_video)
        self.run_btn.clicked.connect(self.toggle_run)
        fullscreen_btn.clicked.connect(self.enter_fullscreen)
        self.ondemand_btn.clicked.connect(self.start_on_demand)

        # ===== 快捷键 =====
        # 删除 Ctrl+Shift+D 快捷键及其相关绑定：不再创建该 QShortcut

        # 空格：暂停/继续（优先点播窗口，否则主窗口）
        s_space = QShortcut(QKeySequence(Qt.Key_Space), self)
        s_space.setContext(Qt.ApplicationShortcut)
        s_space.activated.connect(self.toggle_pause_dispatch)

        # 音量快捷键（全屏也可用）
        s_up = QShortcut(QKeySequence("Ctrl+Up"), self)
        s_up.setContext(Qt.ApplicationShortcut)
        s_up.activated.connect(lambda: self.adjust_volume(+5))

        s_down = QShortcut(QKeySequence("Ctrl+Down"), self)
        s_down.setContext(Qt.ApplicationShortcut)
        s_down.activated.connect(lambda: self.adjust_volume(-5))

        s_m = QShortcut(QKeySequence("Ctrl+M"), self)
        s_m.setContext(Qt.ApplicationShortcut)
        s_m.activated.connect(self.toggle_mute)

        # ===== 左侧控制面板 =====
        left = QWidget()
        left_layout = QVBoxLayout()

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("开始"))
        row1.addWidget(self.start_t)
        row1.addWidget(QLabel("结束"))
        row1.addWidget(self.end_t)
        left_layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(add_seg_btn)
        row2.addWidget(del_seg_btn)
        left_layout.addLayout(row2)

        left_layout.addWidget(QLabel("时间段"))
        left_layout.addWidget(self.segment_box)

        left_layout.addWidget(QLabel("播放列表（可拖动排序，已持久化）"))
        left_layout.addWidget(self.playlist_view, 1)

        row3 = QHBoxLayout()
        row3.addWidget(add_vid_btn)
        row3.addWidget(del_vid_btn)
        left_layout.addLayout(row3)

        left_layout.addWidget(self.graceful_switch_box)

        vol_row = QHBoxLayout()
        vol_row.addWidget(QLabel("音量"))
        vol_row.addWidget(self.volume_slider, 1)
        vol_row.addWidget(mute_btn)
        left_layout.addLayout(vol_row)

        row4 = QHBoxLayout()
        row4.addWidget(self.run_btn)
        row4.addWidget(self.ondemand_btn)
        row4.addWidget(fullscreen_btn)
        left_layout.addLayout(row4)

        left.setLayout(left_layout)

        # ===== Splitter：左控制 + 右视频 =====
        self.splitter = QSplitter(Qt.Horizontal)
        self.left_panel = left
        self.splitter.addWidget(self.left_panel)
        self.splitter.addWidget(self.video)
        self.splitter.setStretchFactor(0, 0)
        self.splitter.setStretchFactor(1, 1)
        self.setCentralWidget(self.splitter)

        # ===== 初始化：加载 CSV =====
        self.load_csv()
        self.refresh_segments()
        self.refresh_playlist_view()

    # =====================
    # 空格暂停/继续：优先点播窗口
    # =====================

    def toggle_pause_dispatch(self):
        if self.ondemand_window is not None and self.ondemand_window.isVisible():
            self.ondemand_window.toggle_pause()
            return
        self.toggle_main_pause()

    def toggle_main_pause(self):
        st = self.player.playbackState()
        if st == QMediaPlayer.PlayingState:
            self.player.pause()
        else:
            self.player.play()

    # =====================
    # 音量控制（主控，联动点播窗口）
    # =====================

    def apply_volume_to_main_audio(self):
        self.audio.setMuted(self.muted)
        if not self.muted:
            self.audio.setVolume(max(0.0, min(1.0, self.volume / 100.0)))

    def set_volume(self, v: int):
        self.volume = max(0, min(100, int(v)))
        self.apply_volume_to_main_audio()
        if self.ondemand_window is not None:
            self.ondemand_window.set_volume_from_main(self.volume, self.muted)

    def adjust_volume(self, delta: int):
        self.volume_slider.setValue(max(0, min(100, self.volume + delta)))

    def toggle_mute(self):
        self.muted = not self.muted
        self.apply_volume_to_main_audio()
        if self.ondemand_window is not None:
            self.ondemand_window.set_volume_from_main(self.volume, self.muted)

    # =====================
    # CSV & 时间段列表
    # =====================

    def load_csv(self):
        self.time_ranges.clear()
        if not Path(CSV_FILE).exists():
            return
        with open(CSV_FILE, newline="", encoding="utf-8") as f:
            for r in csv.DictReader(f):
                try:
                    s = hhmm_to_min(r["start"])
                    e = hhmm_to_min(r["end"])
                    self.time_ranges.append((s, e))
                except Exception:
                    pass
        self.time_ranges.sort()

    def save_csv(self):
        with open(CSV_FILE, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["start", "end"])
            w.writeheader()
            for s, e in self.time_ranges:
                w.writerow({"start": min_to_hhmm(s), "end": min_to_hhmm(e)})

    def refresh_segments(self):
        self.segment_box.clear()
        for s, e in self.time_ranges:
            self.segment_box.addItem(f"{min_to_hhmm(s)} - {min_to_hhmm(e)}")

    # =====================
    # 目录/顺序文件
    # =====================

    def segment_dir(self, seg: tuple[int, int]) -> Path:
        return SEGMENTS_DIR / segment_folder(seg[0], seg[1])

    def order_path(self, seg: tuple[int, int]) -> Path:
        return self.segment_dir(seg) / ORDER_FILE

    def write_order(self, seg: tuple[int, int], names: list[str]):
        p = self.order_path(seg)
        p.write_text("\n".join(names) + ("\n" if names else ""), encoding="utf-8")

    def load_ordered_files(self, seg: tuple[int, int]) -> list[Path]:
        d = self.segment_dir(seg)
        d.mkdir(exist_ok=True)

        files = {p.name: p for p in d.glob("*.mp4")}
        order_file = self.order_path(seg)

        ordered: list[Path] = []
        if order_file.exists():
            names = [line.strip() for line in order_file.read_text(encoding="utf-8").splitlines() if line.strip()]
            for n in names:
                if n in files:
                    ordered.append(files.pop(n))

        for n in sorted(files.keys()):
            ordered.append(files[n])

        self.write_order(seg, [p.name for p in ordered])
        return ordered

    # =====================
    # UI：选中时间段 & 播放列表
    # =====================

    def selected_segment(self):
        idx = self.segment_box.currentIndex()
        if idx < 0 or idx >= len(self.time_ranges):
            return None
        return self.time_ranges[idx]

    def refresh_playlist_view(self):
        self.playlist_view.clear()
        seg = self.selected_segment()
        if not seg:
            return
        ordered = self.load_ordered_files(seg)
        for p in ordered:
            self.playlist_view.addItem(p.name)

    def save_ui_order_for_selected_segment(self):
        seg = self.selected_segment()
        if not seg:
            return
        names = [self.playlist_view.item(i).text() for i in range(self.playlist_view.count())]
        self.write_order(seg, names)

    # =====================
    # 时间段 CRUD
    # =====================

    def add_segment(self):
        s = qtime_to_min(self.start_t.time())
        e = qtime_to_min(self.end_t.time())

        if e < s:
            QMessageBox.warning(self, "非法时间", "结束时间不能早于开始时间")
            return
        if has_overlap(self.time_ranges, (s, e)):
            QMessageBox.warning(self, "时间冲突", "时间段重叠")
            return

        seg = (s, e)
        self.time_ranges.append(seg)
        self.time_ranges.sort()

        d = self.segment_dir(seg)
        d.mkdir(parents=True, exist_ok=True)
        self.write_order(seg, [])

        self.save_csv()
        self.refresh_segments()

    def delete_segment(self):
        seg = self.selected_segment()
        if not seg:
            return

        if self.current_segment == seg:
            self.player.stop()
            self.current_segment = None
            self.current_playlist = []
            self.current_index = 0
            self.waiting_next_segment = None

        self.time_ranges.remove(seg)
        shutil.rmtree(self.segment_dir(seg), ignore_errors=True)

        self.save_csv()
        self.refresh_segments()
        self.playlist_view.clear()

    # =====================
    # 视频 CRUD
    # =====================

    def add_video(self):
        seg = self.selected_segment()
        if not seg:
            QMessageBox.information(self, "提示", "请先选择一个时间段")
            return

        d = self.segment_dir(seg)
        paths, _ = QFileDialog.getOpenFileNames(self, "选择视频", "", "Video (*.mp4 *.mkv *.avi)")
        if not paths:
            return

        for p in paths:
            src = Path(p)
            dst = d / src.name
            if dst.exists():
                continue
            shutil.copy(src, dst)

        self.load_ordered_files(seg)
        self.refresh_playlist_view()

    def delete_video(self):
        seg = self.selected_segment()
        if not seg:
            return
        r = self.playlist_view.currentRow()
        if r < 0:
            return

        name = self.playlist_view.item(r).text()
        fp = self.segment_dir(seg) / name
        if fp.exists() and fp.is_file():
            fp.unlink()

        self.playlist_view.takeItem(r)
        self.save_ui_order_for_selected_segment()

    # =====================
    # 运行按钮：不可运行时不变更 bool 和文字
    # =====================

    def any_segment_has_video(self) -> bool:
        for seg in self.time_ranges:
            if any(self.segment_dir(seg).glob("*.mp4")):
                return True
        return False

    def toggle_run(self):
        if not self.running:
            if not self.time_ranges:
                QMessageBox.warning(self, "无法运行", "没有时间段")
                return
            if not self.any_segment_has_video():
                QMessageBox.warning(self, "无法运行", "所有时间段都没有视频")
                return

        self.running = not self.running
        self.run_btn.setText("结束运行" if self.running else "开始运行")

        if self.running:
            self.clock_timer.start(1000)
            self.check_real_time()
        else:
            self.clock_timer.stop()
            self.player.stop()
            self.current_segment = None
            self.current_playlist = []
            self.current_index = 0
            self.waiting_next_segment = None

    # =====================
    # 全屏：窗口全屏 + 播放画面铺满窗口（隐藏左侧）
    # =====================

    def enter_fullscreen(self):
        self.showFullScreen()
        self.left_panel.hide()

    def exit_fullscreen(self):
        self.showNormal()
        self.left_panel.show()

    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Escape and self.isFullScreen():
            self.exit_fullscreen()
            return
        super().keyPressEvent(e)

    # =====================
    # 调度：真实时间检测 & 切换策略 & 循环播放
    # =====================

    def find_active_segment(self):
        m = now_minutes()
        for seg in self.time_ranges:
            if seg[0] <= m < seg[1]:
                return seg
        return None

    def check_real_time(self):
        if not self.running or self.in_interrupt:
            return

        active = self.find_active_segment()
        if not active:
            return

        if self.current_segment != active:
            if self.current_segment and self.graceful_switch_box.isChecked():
                self.waiting_next_segment = active
            else:
                self.start_segment(active)

    def start_segment(self, seg: tuple[int, int]):
        self.current_segment = seg
        self.waiting_next_segment = None

        self.current_playlist = self.load_ordered_files(seg)
        self.current_index = 0

        if not self.current_playlist:
            return

        self.enter_fullscreen()
        self.play_current()

    def play_current(self):
        if not self.current_playlist:
            return
        self.apply_volume_to_main_audio()
        self.player.setSource(QUrl.fromLocalFile(str(self.current_playlist[self.current_index])))
        self.player.play()

    def on_media_status(self, status):
        if status != QMediaPlayer.EndOfMedia:
            return
        if not self.running or self.in_interrupt:
            return

        if self.waiting_next_segment is not None:
            self.start_segment(self.waiting_next_segment)
            return

        if self.current_playlist:
            self.current_index = (self.current_index + 1) % len(self.current_playlist)
            self.play_current()

    # =====================
    # 点播：打断/恢复（无 Ctrl+Shift+D，仅按钮）
    # =====================

    def start_on_demand(self):
        if self.in_interrupt:
            return

        seg = self.current_segment
        plist = list(self.current_playlist)
        idx = self.current_index
        pos = self.player.position()
        wnext = self.waiting_next_segment
        self.interrupt_state = (seg, plist, idx, pos, wnext)

        self.in_interrupt = True
        self.player.stop()

        self.ondemand_window = OnDemandWindow(
            on_finish=self.resume_from_interrupt,
            get_volume_0_100=lambda: self.volume,
            get_muted=lambda: self.muted
        )
        self.ondemand_window.show()

    def resume_from_interrupt(self, terminated: bool):
        self.in_interrupt = False
        self.ondemand_window = None

        if not self.running:
            return

        if not self.interrupt_state:
            self.check_real_time()
            return

        seg, plist, idx, pos, wnext = self.interrupt_state
        self.interrupt_state = None

        active = self.find_active_segment()

        if seg and active == seg and plist:
            self.current_segment = seg
            self.current_playlist = plist
            self.current_index = min(idx, len(plist) - 1)
            self.waiting_next_segment = wnext

            self.enter_fullscreen()
            self.apply_volume_to_main_audio()
            self.player.setSource(QUrl.fromLocalFile(str(self.current_playlist[self.current_index])))
            self.player.setPosition(pos)
            self.player.play()
            return

        self.current_segment = None
        self.current_playlist = []
        self.current_index = 0
        self.waiting_next_segment = None
        self.check_real_time()


# =====================
# 入口
# =====================

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())
