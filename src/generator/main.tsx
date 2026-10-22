import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "../index.css";
import GeneratorApp from "./GeneratorApp";
import OutputView from "./OutputView";

// One page, two views: the editor, or (opened from the editor with
// ?output=1) the bare output window that follows it.
const params = new URLSearchParams(location.search);
const rootElement = params.get("output") === "1" ? <OutputView screenId={params.get("screen")} /> : <GeneratorApp />;

createRoot(document.getElementById("root")!).render(<StrictMode>{rootElement}</StrictMode>);
