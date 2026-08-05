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

        this.targetContainer.innerHTML = `
            <div style="display: flex; gap: 4px; align-items: center; padding: 2px; width: 100%; box-sizing: border-box;">
                <input type="text" id="exactSearchInput" placeholder="Exact Customer ID..." 
                       style="flex-grow: 1; padding: 4px 6px; border: 1px solid #666; border-radius: 2px; font-size: 12px; outline: none;" />
                <button id="exactSearchBtn" style="padding: 4px 10px; background-color: #0078d4; color: white; border: none; border-radius: 2px; cursor: pointer; font-size: 12px;">Search</button>
                <button id="exactClearBtn" style="padding: 4px 8px; background-color: #f3f2f1; color: #323130; border: 1px solid #ccc; border-radius: 2px; cursor: pointer; font-size: 12px;">Clear</button>
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
    }

    private applyExactFilter(): void {
        const searchValue: string = this.inputElement.value.trim();

        if (!this.dataView || !this.dataView.metadata || !this.dataView.metadata.columns[0]) {
            return;
        }

        const targetColumn = this.dataView.metadata.columns[0];
        const queryName: string = targetColumn.queryName;
        const dotIndex: number = queryName.indexOf(".");

        const target: powerbi.IFilterTarget = {
            table: queryName.substring(0, dotIndex),
            column: queryName.substring(dotIndex + 1)
        };

        if (searchValue === "") {
            this.clearFilter();
            return;
        }

        // Native Power BI Advanced Filter object with strict "Is" equality operator
        const filter: powerbi.IAdvancedFilter = {
            $schema: "http://powerbi.com/product/schema#advanced",
            target: target,
            logicalOperator: "And",
            conditions: [
                {
                    operator: "Is",
                    value: searchValue
                }
            ],
            filterType: powerbi.FilterType.Advanced
        };

        this.host.applyJsonFilter(filter, "general", "filter", FilterAction.merge);
    }

    private clearFilter(): void {
        this.inputElement.value = "";
        this.host.applyJsonFilter(null, "general", "filter", FilterAction.remove);
    }
}
