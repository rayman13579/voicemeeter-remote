import { EventEmitter } from 'events';
import express from 'express';
import voicemeeter, { InterfaceType } from './voicemeeter.mjs';

const app = express();
const eventEmitter = new EventEmitter();

let muted = null;

voicemeeter.init().then(() => {
    voicemeeter.login();
    setInterval(() => {
        if (voicemeeter.isParametersDirty()) {
            sendEventIfMuteChanged();
        }
    }, 500);
});

app.get('/mute-events', (req, res) => {
    res.set({
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Access-Control-Allow-Origin': '*'
    });

    const onMuted = (state) => res.write(`data: ${state}\n\n`);
    eventEmitter.on('muted', onMuted);
    eventEmitter.emit('muted', getMuted());
    req.on('close', () => eventEmitter.removeListener('muted', onMuted));
});

app.listen(8881, () => console.log('Express SSE on :8881'));


const getMuted = () => voicemeeter._getParameterFloat(InterfaceType.strip, 'mute', 0) == 1;

const sendEventIfMuteChanged = () => {
    let newMuted = getMuted();
    if (muted != newMuted) {
        muted = newMuted;
        eventEmitter.emit('muted', muted);
    }
}