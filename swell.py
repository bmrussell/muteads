import argparse
import asyncio
import re
import time
from ctypes import POINTER, cast

from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from winsdk.windows.media.control import (
    GlobalSystemMediaTransportControlsSessionManager as MediaManager,
)

current_volume = None

def get_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    # Get the current volume level
    vol = volume.GetMasterVolumeLevelScalar()
    return vol

def mute_audio():
    global current_volume
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    # Get the current volume level
    current_volume = volume.GetMasterVolumeLevelScalar()

    # Mute the audio
    volume.SetMasterVolumeLevelScalar(0.0, None)

    print(f"Audio muted. Previous volume level: {current_volume}")

def unmute_audio():
    global current_volume
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    # Unmute the audio
    volume.SetMasterVolumeLevelScalar(current_volume, None)

    print(f"Audio unmuted {current_volume}")

    
async def get_media_info(service):
    sessions = await MediaManager.request_async()

    for attempt in range(1,5):                                                          # May fail to get session now & again
        current_session = sessions.get_current_session()
        if current_session:                                                             # there needs to be a media session running
            if re.search(service, current_session.source_app_user_model_id) != None:
                info = await current_session.try_get_media_properties_async()
                # song_attr[0] != '_' ignores system attributes
                info_dict = {song_attr: info.__getattribute__(song_attr) for song_attr in dir(info) if song_attr[0] != '_'}
                # converts winrt vector to list
                info_dict['genres'] = list(info_dict['genres'])

                return info_dict
        time.sleep(5)
    raise Exception('Could not get current media session')


if __name__ == '__main__':
    
    parser = argparse.ArgumentParser()
    parser.add_argument("-s", "--service", type=str, help="Regex for service to mute ads for.", required=True)
    args = parser.parse_args()
    service = args.service
    
    last = None
    current_volume = get_volume()
    while True:
        current_media_info = asyncio.run(get_media_info(service))
        if last != None and last['title'] != current_media_info['title']:
            print(current_media_info)
            if current_media_info['title'] == 'Advertisement':
                mute_audio()
            else:
                unmute_audio()
        time.sleep(2)
        if last != current_media_info:
            last = current_media_info