import * as ExampleComponent from "./Examples";
import * as React from 'react';
import { IPnpReactHookExamplesProps } from './IPnpReactHookExamplesProps';
import { LoadActionOption } from 'pnp-react-hooks/types/options/RenderOptions';
import { PnpHookGlobalOptions, PnpReactOptionProvider } from "pnp-react-hooks";
import { Stack, Nav, INavStyles, INavLink } from '@fluentui/react';

const reactHooksOptions: PnpHookGlobalOptions = {
  disabled: "auto",
  loadActionOption: LoadActionOption.ClearPrevious,
  exception: console.debug
};

const PnpReactHookExamples = (props: IPnpReactHookExamplesProps) =>
{
  const [selectedExample, setExample] = React.useState<{ key: string, component: any }>();
  const componentsLinks: INavLink[] = getComponentNavLinks((key, component) => setExample({ key, component }));
  const ExampleComponentType = selectedExample?.component;

  return (
    <PnpReactOptionProvider value={reactHooksOptions}>
      <Stack horizontal verticalAlign="start" tokens={{ childrenGap: 15 }}>
        <Stack.Item>
          <Nav
            selectedKey={selectedExample?.key}
            styles={navStyles}
            groups={[{
              links: componentsLinks,
              name: "Example Components",
            }]}
          />
        </Stack.Item>
        <Stack.Item style={{ width: "100%" }}>
          {
            ExampleComponentType
              ? <ExampleComponentType />
              : undefined
          }
        </Stack.Item>
      </Stack>
    </PnpReactOptionProvider>
  );
};

export default PnpReactHookExamples;

const getComponentNavLinks = (setter: (key: string, Component: any) => void): INavLink[] =>
{
  return Object.keys(ExampleComponent)
    .sort()
    .map(e => ({
      name: e,
      key: e,
      onClick: () => setter(e, ExampleComponent[e]),
      url: ""
    }));
};

const navStyles: Partial<INavStyles> = {
  root: {
    boxSizing: 'border-box',
    borderRight: '1px solid #333333',
    overflowY: 'auto',
    paddingRight: 5
  },
};