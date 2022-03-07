import * as React from "react";
import { CommandBar, IColumn, ICommandBarItemProps, Icon, IconButton, Link, SearchBox, ShimmeredDetailsList, Stack, Text, Toggle } from "@fluentui/react";
import { useFolderTree, useWebInfo } from "pnp-react-hooks";
import * as dayjs from "dayjs";

interface IExplorerItem
{
    type: "folder" | "file";
    Name: string;
    ServerRelativeUrl: string;
    TimeCreated: string;
    TimeLastModified: string;
}

interface IFolderItem extends IExplorerItem
{
    type: "folder";
    ItemCount: number;
    onClick: () => void;
}

interface IFileItem extends IExplorerItem
{
    type: "file";
    Lenght: number;
    UIVersionLabel: string;
}

export function FileExplorer()
{
    const [searchText, setSearchText] = React.useState<string>();
    const [showHidden, setHidden] = React.useState<boolean>(false);
    const [disable, setDisable] = React.useState<boolean>(false);
    const [refresh, setRefresh] = React.useState({});

    const filter: string = React.useMemo(() =>
    {
        const filterParts: string[] = [];
        if (!showHidden)
        {
            filterParts.push("startswith(Name, '_') eq false");
        }

        if (searchText)
        {
            const normalized = searchText.replace("'", "");
            filterParts.push(`substringof('${normalized}', Name) eq true`);
        }

        return filterParts.length > 0
            ? filterParts.join(" and ")
            : undefined;

    }, [showHidden, searchText]);

    const webInfo = useWebInfo();
    const treeContext = useFolderTree(webInfo?.ServerRelativeUrl, {
        keepPreviousState: false,
        fileQuery: {
            select: ["Name", "Length", "ServerRelativeUrl", "TimeCreated", "TimeLastModified", "UIVersionLabel"],
            filter: filter
        },
        folderFilter: filter,
        disabled: disable ? true : "auto",
    }, [refresh]);

    const _commandItems: ICommandBarItemProps[] = [
        {
            key: 'up',
            text: "Back",
            iconProps: { iconName: 'NavigateBack' },
            disabled: treeContext?.up === undefined,
            onClick: () => treeContext?.up?.()
        },
        {
            key: 'home',
            text: "Home",
            iconProps: { iconName: 'Home' },
            onClick: () => treeContext?.home?.()
        },
        {
            key: 'refresh',
            text: "Refresh",
            iconProps: { iconName: 'Refresh' },
            onClick: () => setRefresh({})
        }
    ];

    const _items: IExplorerItem[] = React.useMemo(() =>
    {
        if (treeContext)
        {
            const folders: IFolderItem[] = treeContext.folders?.map(folder =>
            ({
                type: "folder",
                Name: folder.Name,
                ServerRelativeUrl: folder.ServerRelativeUrl,
                TimeCreated: folder.TimeCreated,
                TimeLastModified: folder.TimeLastModified,
                onClick: () => folder.setAsRoot(),
                ItemCount: folder.ItemCount
            }));

            const files: IFileItem[] = treeContext.files?.map(file =>
            ({
                type: "file",
                Name: file.Name,
                ServerRelativeUrl: file.ServerRelativeUrl,
                TimeCreated: file.TimeCreated,
                TimeLastModified: file.TimeLastModified,
                Lenght: Number(file.Length),
                UIVersionLabel: file.UIVersionLabel,
            }));

            return [].concat(folders, files);
        }
        else
        {
            return undefined;
        }
    }, [treeContext]);

    return (
        <Stack tokens={{ childrenGap: 15 }}>
            <Stack.Item>
                <CommandBar
                    items={_commandItems}
                    ariaLabel="Actions"
                    primaryGroupAriaLabel="Actions"
                />
            </Stack.Item>
            <Stack.Item>
                <Toggle
                    offText="No"
                    onText="Yes"
                    label="Disable Control"
                    checked={disable}
                    onChange={(_, checked) => setDisable(checked)}
                />
            </Stack.Item>
            <Stack.Item>
                <Toggle
                    offText="No"
                    onText="Yes"
                    label="Show Hidden"
                    checked={showHidden}
                    onChange={(_, checked) => setHidden(checked)}
                />
            </Stack.Item>
            <Stack.Item>
                <SearchBox
                    width={300}
                    placeholder="Filter"
                    onSearch={setSearchText}
                    onClear={() => setSearchText(undefined)}
                />
            </Stack.Item>
            <Stack.Item>
                <Text variant="xLargePlus">{treeContext?.root?.Name} </Text>
            </Stack.Item>
            <Stack.Item>
                <ShimmeredDetailsList
                    enableShimmer={!Array.isArray(_items)}
                    items={_items ?? []}
                    columns={_colums}
                />
            </Stack.Item>
        </Stack>
    );
}

const _colums: IColumn[] =
    [
        {
            key: "icon",
            name: "",
            minWidth: 50,
            maxWidth: 50,
            isResizable: true,
            onRender: (item: IExplorerItem) => <Text variant="large">
                <Icon iconName={item.type === "folder" ? "FabricFolderFill" : "Page"} />
            </Text>
        },
        {
            key: "name",
            name: "Name",
            minWidth: 150,
            maxWidth: 250,
            isResizable: true,
            onRender: (item: IFolderItem | IFileItem) =>
                <Link
                    disabled={item.type === "folder" && item.ItemCount >= 5000}
                    value={item.Name}
                    href={item.type === "folder" ? undefined : item.ServerRelativeUrl}
                    onClick={item.type === "folder" ? item.onClick : undefined}
                    target={item.type === "folder" ? undefined : "_blank"}
                >
                    {item.Name}
                </Link >
        },
        {
            key: "copy_url",
            name: "",
            minWidth: 20,
            maxWidth: 20,
            isResizable: false,
            onRender: (item: IExplorerItem) => <IconButton
                iconProps={{ iconName: "Link" }}
                onClick={() => navigator.clipboard.writeText(`${window.location.origin}${item.ServerRelativeUrl}`)
                    .then(e => alert(`Link copied : ${window.location.origin}${item.ServerRelativeUrl}`))
                    .catch(console.error)}
            />
        },
        {
            key: "time_created",
            name: "Created",
            minWidth: 75,
            isResizable: true,
            onRender: (item: IExplorerItem) => dayjs(item?.TimeCreated).format('DD/MM/YYYY')
        },
        {
            key: "time_modified",
            name: "Modified",
            minWidth: 75,
            isResizable: true,
            onRender: (item: IExplorerItem) => dayjs(item?.TimeLastModified).format('DD/MM/YYYY')
        },
        {
            key: "size_info",
            name: "Size",
            minWidth: 100,
            isResizable: true,
            onRender: (item: IFolderItem | IFileItem) => item.type === "file"
                ? `${(item.Lenght / 1024).toLocaleString(undefined, { maximumFractionDigits: 2 })} KB`
                : ""
        }
    ];