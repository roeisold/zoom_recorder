import os
import sys
import json
import time
import schedule
import threading
import datetime
import logging
import cv2
import numpy as np
import pyautogui
import pyaudio
import wave
import subprocess
from screeninfo import get_monitors
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QLabel, QLineEdit, QTimeEdit, QSpinBox, QComboBox, QPushButton, 
                            QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox, QCheckBox)
from PyQt5.QtCore import Qt, QTime, QTimer


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    filename='zoom_recorder.log'
)
logger = logging.getLogger('zoom_recorder')

class ZoomRecorderApp(QMainWindow):
    # GUI implementation remains largely the same as before
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Zoom Meeting Scheduler and Recorder")
        self.setMinimumSize(800, 600)
        
        # Initialize recorder backend
        self.recorder = ZoomMeetingRecorder()
        
        # Create main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        
        # Create form for adding meetings
        form_widget = QWidget()
        form_layout = QVBoxLayout(form_widget)
        
        # Meeting name input
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("Meeting Name:"))
        self.name_input = QLineEdit()
        name_layout.addWidget(self.name_input)
        form_layout.addLayout(name_layout)
        
        # Zoom link input
        link_layout = QHBoxLayout()
        link_layout.addWidget(QLabel("Zoom Link:"))
        self.link_input = QLineEdit()
        link_layout.addWidget(self.link_input)
        form_layout.addLayout(link_layout)
        
        # Time and duration inputs
        time_layout = QHBoxLayout()
        time_layout.addWidget(QLabel("Time:"))
        self.time_input = QTimeEdit()
        self.time_input.setDisplayFormat("HH:mm")
        self.time_input.setTime(QTime(9, 0))
        time_layout.addWidget(self.time_input)
        
        time_layout.addWidget(QLabel("Duration (minutes):"))
        self.duration_input = QSpinBox()
        self.duration_input.setRange(5, 240)
        self.duration_input.setValue(60)
        time_layout.addWidget(self.duration_input)
        form_layout.addLayout(time_layout)
        
        # Day selection
        days_layout = QHBoxLayout()
        days_layout.addWidget(QLabel("Days:"))
        
        self.day_checkboxes = {}
        days = ["Sunday","Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
        for day in days:
            cb = QCheckBox(day)
            self.day_checkboxes[day] = cb
            days_layout.addWidget(cb)
        
        form_layout.addLayout(days_layout)
        
        # Add meeting button
        add_button_layout = QHBoxLayout()
        add_button = QPushButton("Add Meeting")
        add_button.clicked.connect(self.add_meeting)
        add_button_layout.addStretch()
        add_button_layout.addWidget(add_button)
        form_layout.addLayout(add_button_layout)
        
        main_layout.addWidget(form_widget)
        
        # Create table for meetings list
        self.meetings_table = QTableWidget(0, 5)
        self.meetings_table.setHorizontalHeaderLabels(["Meeting Name", "Zoom Link", "Time", "Duration", "Days"])
        self.meetings_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.meetings_table.setSelectionBehavior(QTableWidget.SelectRows)
        
        # Add table buttons
        table_buttons_layout = QHBoxLayout()
        delete_button = QPushButton("Delete Selected")
        delete_button.clicked.connect(self.delete_selected_meeting)
        table_buttons_layout.addWidget(delete_button)
        
        save_button = QPushButton("Save Meetings")
        save_button.clicked.connect(self.save_meetings)
        table_buttons_layout.addWidget(save_button)
        
        main_layout.addWidget(self.meetings_table)
        main_layout.addLayout(table_buttons_layout)
        
        # Status section
        status_layout = QHBoxLayout()
        self.status_label = QLabel("Service status: Stopped")
        status_layout.addWidget(self.status_label)
        
        self.start_service_button = QPushButton("Start Service")
        self.start_service_button.clicked.connect(self.toggle_service)
        self.service_running = False
        status_layout.addWidget(self.start_service_button)
        
        main_layout.addLayout(status_layout)
        
        # Load existing meetings
        self.load_meetings()
        
        # Update timer to check for active recordings
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.update_status)
        self.update_timer.start(5000)  # Update every 5 seconds
    
    def add_meeting(self):
        """Add meeting to the table"""
        name = self.name_input.text()
        link = self.link_input.text()
        time = self.time_input.time().toString("HH:mm")
        duration = self.duration_input.value()
        
        # Validate inputs
        if not name or not link:
            QMessageBox.warning(self, "Input Error", "Meeting name and Zoom link are required.")
            return
        
        if not link.startswith("https://") or "zoom.us" not in link:
            QMessageBox.warning(self, "Input Error", "Please enter a valid Zoom meeting link.")
            return
        
        # Get selected days
        selected_days = []
        for day, checkbox in self.day_checkboxes.items():
            if checkbox.isChecked():
                selected_days.append(day)
        
        if not selected_days:
            QMessageBox.warning(self, "Input Error", "Please select at least one day.")
            return
        
        # Add to table
        row_position = self.meetings_table.rowCount()
        self.meetings_table.insertRow(row_position)
        self.meetings_table.setItem(row_position, 0, QTableWidgetItem(name))
        self.meetings_table.setItem(row_position, 1, QTableWidgetItem(link))
        self.meetings_table.setItem(row_position, 2, QTableWidgetItem(time))
        self.meetings_table.setItem(row_position, 3, QTableWidgetItem(str(duration)))
        self.meetings_table.setItem(row_position, 4, QTableWidgetItem(", ".join(selected_days)))
        
        # Clear form
        self.name_input.clear()
        self.link_input.clear()
        self.time_input.setTime(QTime(9, 0))
        self.duration_input.setValue(60)
        for checkbox in self.day_checkboxes.values():
            checkbox.setChecked(False)
    
    def delete_selected_meeting(self):
        """Delete the selected meeting from the table"""
        selected_rows = set(index.row() for index in self.meetings_table.selectedIndexes())
        for row in sorted(selected_rows, reverse=True):
            self.meetings_table.removeRow(row)
    
    def save_meetings(self):
        """Save meetings to config file"""
        meetings = []
        for row in range(self.meetings_table.rowCount()):
            name = self.meetings_table.item(row, 0).text()
            link = self.meetings_table.item(row, 1).text()
            time = self.meetings_table.item(row, 2).text()
            duration = int(self.meetings_table.item(row, 3).text())
            days = [day.strip() for day in self.meetings_table.item(row, 4).text().split(",")]
            
            meetings.append({
                "name": name,
                "join_url": link,
                "schedule": time,
                "duration_minutes": duration,
                "days": days
            })
        
        # Update config
        self.recorder.config["meetings"] = meetings
        self.recorder.save_config()
        
        QMessageBox.information(self, "Success", "Meetings saved successfully.")
    
    def load_meetings(self):
        """Load meetings from config file"""
        meetings = self.recorder.config.get("meetings", [])
        
        for meeting in meetings:
            row_position = self.meetings_table.rowCount()
            self.meetings_table.insertRow(row_position)
            self.meetings_table.setItem(row_position, 0, QTableWidgetItem(meeting["name"]))
            self.meetings_table.setItem(row_position, 1, QTableWidgetItem(meeting["join_url"]))
            self.meetings_table.setItem(row_position, 2, QTableWidgetItem(meeting["schedule"]))
            self.meetings_table.setItem(row_position, 3, QTableWidgetItem(str(meeting["duration_minutes"])))
            self.meetings_table.setItem(row_position, 4, QTableWidgetItem(", ".join(meeting["days"])))
    
    def toggle_service(self):
        """Start or stop the recording service"""
        if not self.service_running:
            if self.meetings_table.rowCount() == 0:
                QMessageBox.warning(self, "No Meetings", "Please add at least one meeting before starting the service.")
                return
            
            # Save meetings before starting
            self.save_meetings()
            
            # Start recorder service in background thread
            self.service_thread = threading.Thread(target=self.recorder.run_scheduler)
            self.service_thread.daemon = True
            self.service_thread.start()
            
            self.service_running = True
            self.start_service_button.setText("Stop Service")
            self.status_label.setText("Service status: Running")
        else:
            # Stop service
            self.recorder.stop_scheduler()
            self.service_running = False
            self.start_service_button.setText("Start Service")
            self.status_label.setText("Service status: Stopped")
    
    def update_status(self):
        """Update status display with current recorder status"""
        if self.service_running:
            if self.recorder.recording_active:
                current_meeting = self.recorder.current_meeting["name"] if self.recorder.current_meeting else "Unknown"
                self.status_label.setText(f"Service status: Recording meeting '{current_meeting}'")
            else:
                next_meeting = self.recorder.get_next_meeting_info()
                if next_meeting:
                    self.status_label.setText(f"Service status: Waiting for next meeting at {next_meeting['time']} ({next_meeting['name']})")
                else:
                    self.status_label.setText("Service status: Running (no upcoming meetings)")
    
    def closeEvent(self, event):
        """Handle window close event"""
        if self.service_running:
            reply = QMessageBox.question(self, 'Confirm Exit', 
                'The recording service is still running. Stop service and exit?',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                self.recorder.stop_scheduler()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


class ZoomMeetingRecorder:
    def __init__(self, config_file="config.json"):
        """Initialize the Zoom meeting recorder"""
        self.config = self._load_config(config_file)
        self.recordings_path = self.config.get("recordings_path", "recordings")
        
        # Create recordings directory if it doesn't exist
        if not os.path.exists(self.recordings_path):
            os.makedirs(self.recordings_path)
        
        # Track active recordings and scheduler
        self.recording_active = False
        self.scheduler_running = False
        self.current_meeting = None
        
        # Check for FFmpeg
        self.has_ffmpeg = self._check_ffmpeg()
    
    def _check_ffmpeg(self):
        """Check if FFmpeg is available in the system"""
        try:
            subprocess.run(["ffmpeg", "-version"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            logger.info("FFmpeg detected in system")
            return True
        except (subprocess.SubprocessError, FileNotFoundError):
            logger.warning("FFmpeg not found in system PATH. Will save separate audio/video files.")
            return False
    
    def _load_config(self, config_file):
        """Load configuration from JSON file"""
        try:
            with open(config_file, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            logger.error(f"Config file {config_file} not found. Creating a template.")
            default_config = {
                "recordings_path": "recordings",
                "meetings": []
            }
            with open(config_file, 'w') as f:
                json.dump(default_config, f, indent=4)
            return default_config
    
    def save_config(self):
        """Save configuration to file"""
        with open("config.json", 'w') as f:
            json.dump(self.config, f, indent=4)
    
    def join_meeting(self, join_url):
        """Join a Zoom meeting using the join URL"""
        logger.info(f"Joining meeting: {join_url}")
        
        # Open Zoom meeting link using the default browser
        if os.name == 'nt':  # Windows
            os.system(f'start {join_url}')
        elif os.name == 'posix':  # Mac/Linux
            os.system(f'open "{join_url}"')
            
        # Wait for Zoom to launch
        time.sleep(10)
        
        # Look for and click the "Join with Computer Audio" button
        try:
            join_audio_btn = pyautogui.locateOnScreen('join_audio_button.png', confidence=0.8)
            if join_audio_btn:
                pyautogui.click(join_audio_btn)
                logger.info("Clicked 'Join with Computer Audio'")
            else:
                logger.warning("Could not find 'Join with Computer Audio' button")
        except Exception as e:
            logger.error(f"Error when joining audio: {str(e)}")
            
        # Wait for meeting to fully connect
        time.sleep(5)
        
        logger.info("Successfully joined the meeting")
        return True
    
    def start_recording(self, meeting_info):
        """Start screen and audio recording of the meeting"""
        if self.recording_active:
            logger.warning("Recording already in progress")
            return False
            
        meeting_name = meeting_info["name"]
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(self.recordings_path, f"{meeting_name}_{timestamp}")
        
        logger.info(f"Starting recording to {output_file}")
        self.recording_active = True
        self.current_meeting = meeting_info.copy()
        self.current_meeting['output_file'] = output_file
        
        # Start recording in a separate thread
        recording_thread = threading.Thread(
            target=self._record_screen_and_audio,
            args=(output_file,)
        )
        recording_thread.daemon = True
        recording_thread.start()
        
        return True
    
    def _record_screen_and_audio(self, output_file):
        """Record screen and audio directly to MP4 if FFmpeg is available, otherwise save separate files"""
        try:
            # Get screen size
            monitor = get_monitors()[0]
            width, height = monitor.width, monitor.height
            
            # Configure audio capture
            audio_format = pyaudio.paInt16
            channels = 2
            rate = 44100
            chunk = 1024
            
            audio = pyaudio.PyAudio()
            audio_stream = audio.open(format=audio_format, channels=channels,
                                    rate=rate, input=True,
                                    frames_per_buffer=chunk)
            
            audio_frames = []
            video_frames = []
            
            # Set up video writer
            fourcc = cv2.VideoWriter_fourcc(*"XVID")
            fps = 15
            temp_video_file = f"{output_file}_temp.avi"
            video_writer = cv2.VideoWriter(temp_video_file, fourcc, fps, (width, height))
            
            logger.info("Recording started")
            start_time = time.time()
            
            # Record until stopped
            while self.recording_active:
                # Record screen
                screenshot = pyautogui.screenshot()
                frame = np.array(screenshot)
                frame = cv2.cvtColor(frame, cv2.COLOR_RGB2BGR)
                video_writer.write(frame)
                
                # Record audio
                audio_data = audio_stream.read(chunk)
                audio_frames.append(audio_data)
                
            # Cleanup when recording is stopped
            video_writer.release()
            
            # Save audio to WAV file
            temp_audio_file = f"{output_file}_temp.wav"
            wave_file = wave.open(temp_audio_file, 'wb')
            wave_file.setnchannels(channels)
            wave_file.setsampwidth(audio.get_sample_size(audio_format))
            wave_file.setframerate(rate)
            wave_file.writeframes(b''.join(audio_frames))
            wave_file.close()
            
            # Clean up audio resources
            audio_stream.stop_stream()
            audio_stream.close()
            audio.terminate()
            
            # Combine audio and video using FFmpeg if available
            if self.has_ffmpeg:
                mp4_file = f"{output_file}.mp4"
                cmd = [
                    "ffmpeg", "-y",
                    "-i", temp_video_file,
                    "-i", temp_audio_file,
                    "-c:v", "libx264",
                    "-c:a", "aac",
                    "-strict", "experimental",
                    mp4_file
                ]
                logger.info(f"Combining audio and video with FFmpeg: {' '.join(cmd)}")
                result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                
                if result.returncode == 0:
                    logger.info(f"Successfully created MP4 file: {mp4_file}")
                    # Remove temp files
                    os.remove(temp_video_file)
                    os.remove(temp_audio_file)
                else:
                    logger.error(f"Error combining files with FFmpeg: {result.stderr.decode()}")
                    logger.info(f"Keeping separate audio and video files.")
            else:
                # If FFmpeg not available, rename the temp files to final names
                final_video = f"{output_file}.avi"
                final_audio = f"{output_file}.wav"
                os.rename(temp_video_file, final_video)
                os.rename(temp_audio_file, final_audio)
                logger.info(f"Video saved to {final_video}")
                logger.info(f"Audio saved to {final_audio}")
            
            # Calculate duration
            duration = time.time() - start_time
            logger.info(f"Recording finished. Duration: {duration:.2f} seconds")
            
        except Exception as e:
            logger.error(f"Error during recording: {str(e)}")
    
    def stop_recording(self):
        """Stop the current recording"""
        if not self.recording_active:
            logger.warning("No active recording to stop")
            return False
            
        self.recording_active = False
        logger.info("Stopping recording...")
        
        # Wait for resources to be released
        time.sleep(5)
        
        self.current_meeting = None
        return True
    
    def leave_meeting(self):
        """Leave the current Zoom meeting by closing the Zoom process"""
        logger.info("Leaving meeting by closing Zoom")
        
        try:
            # Kill Zoom process
            if os.name == 'nt':  # Windows
                os.system('taskkill /f /im Zoom.exe')
            elif os.name == 'posix':  # Mac/Linux
                os.system('pkill -f zoom')
                
            logger.info("Zoom process terminated")
            
            # Wait a moment for process to fully terminate
            time.sleep(3)
            return True
            
        except Exception as e:
            logger.error(f"Error when terminating Zoom: {str(e)}")
            return False
    
    def execute_scheduled_task(self, meeting_info):
        """Execute the full meeting join, record, and leave workflow"""
        logger.info(f"Executing scheduled task for meeting {meeting_info['name']}")
        
        try:
            # Join the meeting
            success = self.join_meeting(meeting_info["join_url"])
            if not success:
                logger.error("Failed to join meeting")
                return False
                
            # Start recording
            self.start_recording(meeting_info)
            
            # Record for the duration of the meeting
            meeting_duration = int(meeting_info["duration_minutes"])
            logger.info(f"Recording for {meeting_duration} minutes")
            
            # Wait for the meeting duration
            time.sleep(meeting_duration * 60)
            
            # Stop recording and leave meeting
            self.stop_recording()
            self.leave_meeting()
            
            logger.info(f"Completed scheduled task for meeting {meeting_info['name']}")
            return True
            
        except Exception as e:
            logger.error(f"Error executing scheduled task: {str(e)}")
            # Try to clean up
            try:
                self.stop_recording()
                self.leave_meeting()
            except:
                pass
            return False
    
    def get_next_meeting_info(self):
        """Get information about the next scheduled meeting"""
        meetings = self.config.get("meetings", [])
        if not meetings:
            return None
        
        now = datetime.datetime.now()
        current_day = now.strftime("%A")
        current_time_str = now.strftime("%H:%M")
        
        next_meeting = None
        soonest_time_diff = float('inf')
        
        for meeting in meetings:
            if current_day in meeting["days"]:
                meeting_time = datetime.datetime.strptime(meeting["schedule"], "%H:%M").time()
                meeting_datetime = datetime.datetime.combine(now.date(), meeting_time)
                
                # If meeting is today but already passed, skip it
                if meeting_datetime < now:
                    continue
                
                time_diff = (meeting_datetime - now).total_seconds()
                if time_diff < soonest_time_diff:
                    soonest_time_diff = time_diff
                    next_meeting = meeting.copy()
                    next_meeting["time"] = meeting["schedule"]
        
        # Check tomorrow if no meetings found today
        if not next_meeting:
            tomorrow = now + datetime.timedelta(days=1)
            tomorrow_day = tomorrow.strftime("%A")
            
            for meeting in meetings:
                if tomorrow_day in meeting["days"]:
                    meeting_time = meeting["schedule"]
                    next_meeting = meeting.copy()
                    next_meeting["time"] = f"Tomorrow at {meeting_time}"
                    break
        
        return next_meeting
    
    def run_scheduler(self):
        """Run the scheduler to execute meetings"""
        self.scheduler_running = True
        logger.info("Starting scheduler")
        
        # Clear any existing jobs
        schedule.clear()
        
        # Schedule all meetings
        for meeting in self.config.get("meetings", []):
            schedule_time = meeting["schedule"]
            days = meeting["days"]
            
            # Create a job for each day
            for day in days:
                job = schedule.every()
                
                if day == "Monday":
                    job = job.monday
                elif day == "Tuesday":
                    job = job.tuesday
                elif day == "Wednesday":
                    job = job.wednesday
                elif day == "Thursday":
                    job = job.thursday
                elif day == "Friday":
                    job = job.friday
                elif day == "Saturday":
                    job = job.saturday
                elif day == "Sunday":
                    job = job.sunday
                
                job.at(schedule_time).do(self.execute_scheduled_task, meeting)
                logger.info(f"Scheduled meeting '{meeting['name']}' for {day} at {schedule_time}")
        
        # Run pending jobs in a loop
        while self.scheduler_running:
            schedule.run_pending()
            time.sleep(1)
        
        logger.info("Scheduler stopped")
    
    def stop_scheduler(self):
        """Stop the scheduler"""
        self.scheduler_running = False
        
        # Stop any active recording
        if self.recording_active:
            self.stop_recording()
            self.leave_meeting()
        
        # Clear scheduled jobs
        schedule.clear()
        logger.info("Scheduler and all jobs cleared")


if __name__ == '__main__':
    # Create and show app
    app = QApplication(sys.argv)
    window = ZoomRecorderApp()
    window.show()
    sys.exit(app.exec_())