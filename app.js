// Freeze Frame: a Teams video filter that holds the last camera frame.
//
// Choosing the "Freeze" filter in Teams' Video effects captures the next
// frame and keeps sending it. Choosing "None" (or "Go live" on this page)
// switches back to the live camera. Your camera never turns off.

const { app, videoEffects } = window.microsoftTeams;

// Must match meetingExtensionDefinition.videoFilters[].id in the manifest.
const FREEZE_EFFECT_ID = "b5067bc7-807a-4b44-b301-2c18467c00a3";

let frozen = false;
let heldFrame = null; // WebCodecs VideoFrame, for hosts that use videoFrameHandler
let heldBuffer = null; // NV12 bytes, for hosts that use videoBufferHandler

const statusEl = document.getElementById("status");
const toggleEl = document.getElementById("toggle");
const logEl = document.getElementById("log");

function log(message) {
  console.log("[freeze]", message);
  logEl.textContent = `${new Date().toLocaleTimeString()}  ${message}\n` + logEl.textContent;
}

function setFrozen(value) {
  frozen = value;
  if (!frozen) {
    heldFrame?.close();
    heldFrame = null;
    heldBuffer = null;
  }
  statusEl.textContent = frozen ? "Frozen" : "Live";
  statusEl.dataset.state = frozen ? "frozen" : "live";
  toggleEl.textContent = frozen ? "Go live" : "Freeze";
}

// Host passes WebCodecs VideoFrames and expects one back.
async function videoFrameHandler({ videoFrame }) {
  if (!frozen) return videoFrame;

  if (!heldFrame) {
    heldFrame = videoFrame.clone();
    return videoFrame;
  }

  // Re-stamp the held image with the incoming timestamp so the pipeline
  // keeps treating it as a fresh frame.
  const out = new VideoFrame(heldFrame, { timestamp: videoFrame.timestamp });
  videoFrame.close();
  return out;
}

// Host passes an NV12 buffer that we modify in place.
function videoBufferHandler(bufferData, notifyVideoFrameProcessed, notifyError) {
  try {
    if (frozen) {
      const buf = bufferData.videoFrameBuffer;
      if (!heldBuffer || heldBuffer.length !== buf.length) {
        heldBuffer = new Uint8ClampedArray(buf); // copy
      } else {
        buf.set(heldBuffer);
      }
    }
    notifyVideoFrameProcessed();
  } catch (e) {
    notifyError(String(e));
  }
}

// Teams calls this when the user picks or clears our filter in Video effects.
async function onEffectChanged(effectId) {
  log(`Teams applied effect: ${effectId ?? "none"}`);
  if (effectId && effectId !== FREEZE_EFFECT_ID) {
    throw videoEffects.EffectFailureReason.InvalidEffectId;
  }
  setFrozen(effectId === FREEZE_EFFECT_ID);
}

toggleEl.addEventListener("click", () => {
  // Ask Teams to apply/clear the effect; it calls onEffectChanged back.
  // In a meeting Teams may also need you to click its Apply button.
  if (frozen) {
    videoEffects.notifySelectedVideoEffectChanged(videoEffects.EffectChangeType.EffectDisabled, undefined);
  } else {
    videoEffects.notifySelectedVideoEffectChanged(videoEffects.EffectChangeType.EffectChanged, FREEZE_EFFECT_ID);
  }
});

async function start() {
  setFrozen(false);
  try {
    await app.initialize();
  } catch (e) {
    log(`Not running inside Teams (${e.message ?? e}). Open this from Teams' Video effects.`);
    return;
  }
  const ctx = await app.getContext();
  log(`Teams host: ${ctx.app?.host?.clientType ?? "?"} / ${ctx.app?.host?.name ?? "?"}, frame context: ${ctx.page?.frameContext ?? "?"}`);

  if (!videoEffects.isSupported()) {
    log("videoEffects.isSupported() = false: this Teams client/tenant doesn't expose the video API to this app.");
    toggleEl.disabled = true;
    return;
  }

  videoEffects.registerForVideoEffect(onEffectChanged);
  videoEffects.registerForVideoFrame({
    videoFrameHandler,
    videoBufferHandler,
    config: { format: videoEffects.VideoFrameFormat.NV12 },
  });
  log("Registered for video frames. Pick the Freeze filter to hold the picture.");
}

start();
