import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import FilterAction = powerbi.FilterAction;

export class Visual implements IVisual {
    private targetContainer: HTMLElement;
    private inputElement: HTMLInputElement;
    private searchButton: HTMLButtonElement;
    private clearButton: HTMLButtonElement;
    private host: IVisualHost;
    private dataView: powerbi.DataView;

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
            this.applyExactFilter();
        });
        this.clearButton.addEventListener("click", () => {
            this.clearFilter();
        });
        this.inputElement.addEventListener("keydown", (e: KeyboardEvent) => {
            if (e.key === "Enter") {
                this.applyExactFilter();
            }
        });
    }

    public update(options: VisualUpdateOptions): void {
        if (options.dataViews && options.dataViews[0]) {
            this.dataView = options.dataViews[0];
        }

        // Dynamic scaling based on visual container height & width
        if (options.viewport) {
            const height = options.viewport.height;
            const width = options.viewport.width;

            let fontSize = Math.max(9, Math.min(14, Math.floor(height * 0.35)));
            if (width < 140) {
                fontSize = Math.max(8, fontSize - 2);
            }

            const btnPadding = Math.max(2, Math.min(6, Math.floor(height * 0.12)));
            const iconSize = Math.max(10, Math.min(16, Math.floor(height * 0.38)));

            this.inputElement.style.fontSize = `${fontSize}px`;
            this.searchButton.style.padding = `${btnPadding}px`;
            this.clearButton.style.padding = `${btnPadding}px`;

            const svgs = this.targetContainer.querySelectorAll("svg");
            svgs.forEach((svg) => {
                svg.setAttribute("width", `${iconSize}`);
                svg.setAttribute("height", `${iconSize}`);
            });
        }
    }

    private applyExactFilter(): void {
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

        const jsonFilter = {
            $schema: "http://powerbi.com/product/schema#advanced",
            target: target,
            logicalOperator: "And",
            conditions: [
                {
                    operator: "Is",
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
}
