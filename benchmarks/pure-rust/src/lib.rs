//! Code shared by the pure Rust benchmark binaries.

pub mod export;

/// Peak resident set size of this process so far, in bytes.
#[cfg(unix)]
pub fn max_rss_bytes() -> Option<u64> {
    let mut usage = std::mem::MaybeUninit::<libc::rusage>::zeroed();
    let status = unsafe { libc::getrusage(libc::RUSAGE_SELF, usage.as_mut_ptr()) };
    if status != 0 {
        return None;
    }

    let max_rss = unsafe { usage.assume_init() }.ru_maxrss;
    if max_rss < 0 {
        return None;
    }

    let max_rss = max_rss as u64;
    if cfg!(any(target_os = "macos", target_os = "ios")) {
        Some(max_rss)
    } else {
        Some(max_rss.saturating_mul(1_024))
    }
}

/// Peak resident set size of this process so far, in bytes.
#[cfg(not(unix))]
pub fn max_rss_bytes() -> Option<u64> {
    None
}
