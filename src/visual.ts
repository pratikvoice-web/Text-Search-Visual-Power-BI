import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import FilterAction = powerbi.FilterAction;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstance = powerbi.VisualObjectInstance;
import VisualObjectInstanceEnumerationObject = powerbi.VisualObjectInstanceEnumerationObject;

export class Visual implements IVisual {
    private targetContainer: HTMLElement;
    private inputElement: HTMLInputElement;
    private searchButton: HTMLButtonElement;
    private clearButton: HTMLButtonElement;
    private host: IVisualHost;
    private dataView: powerbi.DataView;

    // Formatting Settings Defaults
    private customFontSize: number = 0; // 0 indicates automatic responsive scaling
    private matchType: string = "Is";   // "Is" (Exact) or "Contains" (Fuzzy/Substring)

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.targetContainer = options.element;

        this.targetContainer.style.width = "100%";
        this.targetContainer.style.height = "100%";
        this.targetContainer.style.overflow = "hidden";
        this.targetContainer.style.boxSizing = "border-box";

        const searchIconSvg = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>`;
        const clearIconSvg = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>`;

        this.targetContainer.innerHTML = `
            <div id="searchVisualContainer" style="display: flex; gap: 4px; align-items: center; padding: 2px 4px; width: 100%; height: 100%; box-sizing: border-box; overflow: hidden;">
                <input type="text" id="exactSearchInput" placeholder="Enter text to search..." 
                       style="flex: 1 1 auto; min-width: 0; width: 100%; padding: 4px 6px; border: 1px solid #666; border-radius: 3px; font-size: 12px; outline: none; box-sizing: border-box;" />
                <button id="exactSearchBtn" title="Search" style="flex: 0 0 auto; display: inline-flex; align-items: center; justify-content: center; padding: 5px; background-color: #0078d4; color: white; border: none; border-radius: 3px; cursor: pointer;">
                    ${searchIconSvg}
                </button>
                <button id="exactClearBtn" title="Clear" style="flex: 0 0 auto; display: inline-flex; align-items: center; justify-content: center; padding: 5px; background-color: #f3f2f1; color: #323130; border: 1px solid #ccc; border-radius: 3px; cursor: pointer;">
                    ${clearIconSvg}
                </button>
            </div>
        `;

        this.inputElement = this.targetContainer.querySelector("#exactSearchInput") as HTMLInputElement;
        this.searchButton = this.targetContainer.querySelector("#exactSearchBtn") as HTMLButtonElement;
        this.clearButton = this.targetContainer.querySelector("#exactClearBtn") as HTMLButtonElement;

        this.searchButton.addEventListener("click", () => {
            this.applyFilter();
        });
        this.clearButton.addEventListener("click", () => {
            this.clearFilter();
        });
        this.inputElement.addEventListener("keydown", (e: KeyboardEvent) => {
            if (e.key === "Enter") {
                this.applyFilter();
            }
        });
    }

    public update(options: VisualUpdateOptions): void {
        if (options.dataViews && options.dataViews[0]) {
            this.dataView = options.dataViews[0];

            // Parse custom formatting properties
            const objects = this.dataView.metadata?.objects;
            if (objects && objects["searchSettings"]) {
                const searchObj = objects["searchSettings"];
                if (searchObj["fontSize"] !== undefined) {
                    this.customFontSize = searchObj["fontSize"] as number;
                }
                if (searchObj["matchType"] !== undefined) {
                    this.matchType = searchObj["matchType"] as string;
                }
            }
        }

        // Synchronize incoming filter state from synced slicers or bookmarks
        if (options.jsonFilters && options.jsonFilters.length > 0) {
            const rawFilter = options.jsonFilters[0] as any;
            if (rawFilter && rawFilter.conditions && rawFilter.conditions.length > 0) {
                const conditionVal = rawFilter.conditions[0].value;
                if (conditionVal !== undefined && conditionVal !== null) {
                    this.inputElement.value = String(conditionVal);
                }
            }
        } else if (options.jsonFilters && options.jsonFilters.length === 0) {
            // Filter was cleared on another synced page
            if (this.inputElement.value !== "") {
                this.inputElement.value = "";
            }
        }

        // Apply sizing logic (Custom Font Size vs Dynamic Viewport Scaling)
        if (options.viewport) {
            const height = options.viewport.height;
            const width = options.viewport.width;

            let computedFontSize: number;
            if (this.customFontSize && this.customFontSize > 0) {
                computedFontSize = this.customFontSize;
                this.inputElement.style.fontSize = `${computedFontSize}pt`;
            } else {
                computedFontSize = Math.max(9, Math.min(14, Math.floor(height * 0.35)));
                if (width < 140) {
                    computedFontSize = Math.max(8, computedFontSize - 2);
                }
                this.inputElement.style.fontSize = `${computedFontSize}px`;
            }

            const btnPadding = Math.max(2, Math.min(6, Math.floor(height * 0.12)));
            const iconSize = Math.max(10, Math.min(16, Math.floor(height * 0.38)));

            this.searchButton.style.padding = `${btnPadding}px`;
            this.clearButton.style.padding = `${btnPadding}px`;

            const svgs = this.targetContainer.querySelectorAll("svg");
            svgs.forEach((svg) => {
                svg.setAttribute("width", `${iconSize}`);
                svg.setAttribute("height", `${iconSize}`);
            });
        }
    }

    private applyFilter(): void {
        const searchValue: string = this.inputElement.value.trim();

        if (!this.dataView || !this.dataView.metadata || !this.dataView.metadata.columns[0]) {
            return;
        }

        const targetColumn = this.dataView.metadata.columns[0];
        const queryName: string = targetColumn.queryName;
        const dotIndex: number = queryName.indexOf(".");

        if (dotIndex === -1) {
            return;
        }

        const target = {
            table: queryName.substring(0, dotIndex),
            column: queryName.substring(dotIndex + 1)
        };

        if (searchValue === "") {
            this.clearFilter();
            return;
        }

        // Dynamically select filter operator based on format pane setting: "Is" or "Contains"
        const selectedOperator = (this.matchType === "Contains") ? "Contains" : "Is";

        const jsonFilter = {
            $schema: "http://powerbi.com/product/schema#advanced",
            target: target,
            logicalOperator: "And",
            conditions: [
                {
                    operator: selectedOperator,
                    value: searchValue
                }
            ]
        };

        this.host.applyJsonFilter(
            jsonFilter as unknown as powerbi.IFilter,
            "general",
            "filter",
            FilterAction.merge
        );
    }

    private clearFilter(): void {
        this.inputElement.value = "";
        this.host.applyJsonFilter(null, "general", "filter", FilterAction.remove);
    }

    // Expose searchSettings to the Power BI Format Visual Pane
    public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] | VisualObjectInstanceEnumerationObject {
        const instances: VisualObjectInstance[] = [];

        if (options.objectName === "searchSettings") {
            instances.push({
                objectName: options.objectName,
                properties: {
                    fontSize: this.customFontSize,
                    matchType: this.matchType
                },
                selector: null
            });
        }

        return instances;
    }
}
