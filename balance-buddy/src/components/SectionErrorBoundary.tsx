import { Component, type ReactNode } from "react";
import { reportLovableError } from "../lib/lovable-error-reporting";

type Props = {
  /** Short human label for the section, shown in the fallback. */
  title?: string;
  children: ReactNode;
};

type State = { error: Error | null };

/**
 * Section-level error boundary.
 *
 * Wraps an independent piece of the page (a table, a chart panel, the detail
 * pane…) so a render crash inside it shows a small, recoverable message INSTEAD
 * of tearing down the whole app and losing the user's uploaded data. The rest of
 * the page — upload zone, summary, other sections — keeps working.
 *
 * A self-contained class boundary (React has no hook equivalent) avoids adding a
 * dependency. `resetKey` lets a parent clear the error when inputs change.
 */
export class SectionErrorBoundary extends Component<Props & { resetKey?: unknown }, State> {
  state: State = { error: null };

  static getDerivedStateFromError(error: Error): State {
    return { error };
  }

  componentDidCatch(error: Error) {
    // Keep the existing telemetry path; never let reporting itself throw.
    try {
      reportLovableError(error, { boundary: `section:${this.props.title ?? "unknown"}` });
    } catch {
      /* ignore */
    }
    console.error(`[SectionErrorBoundary:${this.props.title ?? "?"}]`, error);
  }

  componentDidUpdate(prev: Props & { resetKey?: unknown }) {
    // Auto-recover when the parent signals the inputs changed (e.g. a new run).
    if (this.state.error && prev.resetKey !== this.props.resetKey) {
      this.setState({ error: null });
    }
  }

  render() {
    if (this.state.error) {
      return (
        <div className="rounded-2xl border border-rose-200 bg-rose-50/70 p-5 text-center">
          <div className="text-sm font-black text-rose-700">
            {this.props.title ?? "This section"} hit a snag
          </div>
          <p className="mx-auto mt-1 max-w-md text-[11px] font-medium text-rose-600/90">
            This part of the page couldn't render, but your uploaded files and the rest of the
            results are safe. Try this section again, or switch to another tab.
          </p>
          <pre className="mx-auto mt-2 max-w-md overflow-auto rounded-lg bg-white/70 px-2 py-1 text-left text-[9px] text-rose-500">
            {this.state.error.message}
          </pre>
          <button
            onClick={() => this.setState({ error: null })}
            className="mt-3 inline-flex items-center justify-center rounded-lg bg-rose-600 px-3 py-1.5 text-[11px] font-bold text-white transition-colors hover:bg-rose-700"
          >
            Try again
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}
