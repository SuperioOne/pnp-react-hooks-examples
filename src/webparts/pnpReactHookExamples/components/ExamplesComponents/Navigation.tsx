import { CommandBar, ICommandBarItemProps, INavLink, Nav, Separator, Shimmer, Stack, Text } from "@fluentui/react";
import { INavNodeInfo } from "@pnp/sp/navigation";
import { useNavigation } from "pnp-react-hooks";
import * as React from "react";

export function Navigation()
{
    return (
        <Stack tokens={{ childrenGap: 15 }}>
            <Stack.Item>
                <Text variant="xLargePlus">Quicklaunch Navigation</Text>
                <Separator />
            </Stack.Item>
            <Stack.Item>
                <QuickLaunch />
            </Stack.Item>
            <Stack.Item>
                <Text variant="xLargePlus"> Navigation</Text>
                <Separator />
            </Stack.Item>
            <Stack.Item>
                <TopNav />
            </Stack.Item>
        </Stack>
    );
}

function TopNav()
{
    const topNav = useNavigation({
        type: "topNavigation"
    });

    const menuItems: ICommandBarItemProps[] = React.useMemo(() =>
    {
        return topNav?.map((e): ICommandBarItemProps =>
        ({
            key: e.Id.toString(),
            text: e.Title,
            href: e.Url
        }));

    }, [topNav]);

    return (
        <Shimmer isDataLoaded={menuItems !== undefined} >
            <CommandBar
                items={menuItems ?? []}
                ariaLabel="Top Navigation"
                primaryGroupAriaLabel="Top navigation links"
            />
        </Shimmer>
    );
}

function QuickLaunch()
{
    const quickLaunch = useNavigation({
        type: "quickLaunch",
        query: {
            expand: ["children"]
        }
    }) as INavNodeWithChild[];

    const menuItems = React.useMemo(() =>
    {
        return [
            {
                links: quickLaunch?.map((e): INavLink =>
                ({
                    name: e.Title,
                    links: _getLinks(e),
                    url: e.Url
                }))
            }
        ];
    }, [quickLaunch]);

    return (
        <Shimmer isDataLoaded={menuItems !== undefined} >
            <Nav
                groups={menuItems ?? []}
                ariaLabel="Quicklaunch"
            />
        </Shimmer>
    );
}

interface INavNodeWithChild extends INavNodeInfo
{
    Children?: INavNodeWithChild[];
}

function _getLinks(node: INavNodeWithChild): INavLink[]
{
    if (node?.Children)
    {
        return node.Children.map(e => ({
            name: e.Title,
            url: e.Url
        }));
    }
    else
    {
        return undefined;
    }
}