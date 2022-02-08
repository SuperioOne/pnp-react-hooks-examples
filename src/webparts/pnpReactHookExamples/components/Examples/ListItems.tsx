import * as React from "react";
import { DetailsList, DetailsListLayoutMode, Dropdown, IColumn, IDropdownOption, IStackItemStyles, IStackStyles, Shimmer, Slider, Stack } from "@fluentui/react";
import { useListItems, useLists } from "pnp-react-hooks";

export function ListItems()
{
    const rangeRef = React.useRef<number>();
    const [sliderValue, setSliderValue] = React.useState(10);
    const [sliderLowerValue, setSliderLowerValue] = React.useState(0);
    const [range, setRange] = React.useState<[number, number]>([sliderLowerValue, sliderValue]);
    const [selectedList, setSelectedList] = React.useState<string>(null);

    // get lists from current web
    const listsOptions: IDropdownOption[] = useLists({
        query: {
            select: ["Title", "Id"],
            orderBy: "Title",
            orderyByAscending: true
        },
    })?.map(e => ({ key: e.Id, text: e.Title }));

    const items = useListItems(selectedList, {
        query: {
            select: ["Title", "Id", "Author/Title", "Editor/Title", "Modified", "Created"],
            expand: ["Author", "Editor"],
            skip: range[0],
            top: range[1]-range[0],
        },
    });

    const onChange = (_: unknown, newRange: [number, number]) =>
    {
        if (rangeRef.current > 0) clearTimeout(rangeRef.current);

        setSliderLowerValue(newRange[0]);
        setSliderValue(newRange[1]);

        rangeRef.current = setTimeout(() =>
        {
            setRange(newRange);

            if (rangeRef.current > 0) clearTimeout(rangeRef.current);
        }, 1000);
    };

    return (
        <Stack tokens={{ childrenGap: 25 }}>
            <Stack.Item>
                <Dropdown
                    options={listsOptions}
                    disabled={!Array.isArray(listsOptions)}
                    styles={dropdownStyles}
                    label="Lists"
                    placeholder="Select a list"
                    selectedKey={selectedList}
                    onChange={(_, opt) =>
                    {
                        setSelectedList(opt.key?.toString());
                    }}
                />
            </Stack.Item>
            <Stack.Item styles={stackStyles}>
                <Slider
                    label="Item Range"
                    lowerValue={sliderLowerValue}
                    max={500}
                    min={0}
                    step={1}
                    onChange={onChange}
                    ranged
                    showValue
                    value={sliderValue}
                />
            </Stack.Item>
            <Stack.Item>
                {
                    selectedList ?
                        <Shimmer isDataLoaded={items?.length > 0}>
                            <DetailsList
                                items={items ?? []}
                                columns={colums}
                                layoutMode={DetailsListLayoutMode.justified}
                            />
                        </Shimmer>
                        : undefined
                }
            </Stack.Item>
        </Stack>
    );
}

const stackStyles: Partial<IStackItemStyles> = { root: { maxWidth: 300 } };
const colums: IColumn[] =
    [
        {
            key: "Id",
            name: "Id",
            minWidth: 20,
            maxWidth: 50,
            isResizable: true,
            fieldName: "Id"
        },
        {
            key: "Title",
            name: "Title",
            minWidth: 100,
            isResizable: true,
            fieldName: "Title"
        },
        {
            key: "Created By",
            name: "Created By",
            minWidth: 100,
            isResizable: true,
            onRender: item => item?.Author?.Title
        },
        {
            key: "Modified By",
            name: "Modified By",
            minWidth: 100,
            isResizable: true,
            onRender: item => item?.Editor?.Title
        },
        {
            key: "Created",
            name: "Created",
            minWidth: 100,
            isResizable: true,
            fieldName: "Created"
        },
        {
            key: "Modified",
            name: "Modified",
            minWidth: 100,
            isResizable: true,
            fieldName: "Modified"
        }

    ];

const dropdownStyles = { dropdown: { width: 300 } };