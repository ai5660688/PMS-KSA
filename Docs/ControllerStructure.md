# HomeController Partial Layout

The legacy `HomeController.WeldersAndLogs.cs` file now exists only as an empty placeholder to avoid reintroducing duplicate endpoint definitions. The controller is organized into focused partials:

- `PMS/Controllers/HomeController.Welders.List.cs` – Welder list UI, delete endpoint, Excel export, and qualification ZIP download.
- `PMS/Controllers/HomeController.Welders.Save.cs` – Welder edit/add flows plus qualification CRUD and helper APIs.
- `PMS/Controllers/HomeController.Projects.cs` – Project log CRUD and export endpoints.
- `PMS/Controllers/HomeController.Wps.cs` – WPS log CRUD and export endpoints.
- Additional feature areas (e.g., DWR, Daily Fit-up) live in their own `HomeController.*.cs` partials.

If you need to add a new action, place it in the partial that matches the feature area (or create a new partial with a clear name). Do **not** recreate a combined controller file—doing so will immediately generate `CS0111` duplicate-member errors.
