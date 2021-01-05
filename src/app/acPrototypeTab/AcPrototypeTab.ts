import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/acPrototypeTab/index.html")
@PreventIframe("/acPrototypeTab/login.html")

export class AcPrototypeTab {
}
