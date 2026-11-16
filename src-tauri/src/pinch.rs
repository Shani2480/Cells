//! Touchpad pinch on Linux. WebKitGTK takes a pinch as page zoom, magnifying the whole window like a web page; the
//! page never hears of it. A zoom gesture on the webview, in the capture phase, claims the pinch first (so WebKit's
//! own gesture is cancelled) and hands the scale to the page, which zooms the sheet (`window.__cellsPinch`).

use gtk::prelude::*;
use webkit2gtk::WebViewExt;

pub fn install(window: &tauri::WebviewWindow) {
    let _ = window.with_webview(|platform| {
        let view: webkit2gtk::WebView = platform.inner();
        let gesture = gtk::GestureZoom::new(&view);
        gesture.set_propagation_phase(gtk::PropagationPhase::Capture);
        let v = view.clone();
        gesture.connect_begin(move |g, _| {
            g.set_state(gtk::EventSequenceState::Claimed);
            send(&v, "begin", 1.0);
        });
        let v = view.clone();
        gesture.connect_scale_changed(move |_, scale| send(&v, "change", scale));
        let v = view.clone();
        gesture.connect_end(move |_, _| send(&v, "end", 1.0));
        // A GTK 3 gesture lives only as long as a reference to it: the view keeps it.
        unsafe { view.set_data("cells-pinch", gesture) };
    });
}

fn send(view: &webkit2gtk::WebView, phase: &str, scale: f64) {
    let script = format!("window.__cellsPinch && window.__cellsPinch('{phase}', {scale})");
    view.evaluate_javascript(&script, None, None, None::<&gtk::gio::Cancellable>, |_| {});
}
