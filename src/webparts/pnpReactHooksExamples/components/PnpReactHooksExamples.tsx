import * as React from 'react';
import ExampleComponents from './ExampleComponents';
import type { IPnpReactHooksExamplesProps } from './IPnpReactHooksExamplesProps';
import { Stack, Nav, INavStyles, INavLink } from '@fluentui/react';
import { PnpHookOptionProvider } from 'pnp-react-hooks';

const getComponentNavLinks = (setter: (key: string, Component: React.FunctionComponent) => void): INavLink[] => {
  return ExampleComponents.map(e => ({
    name: e[0],
    key: e[0],
    onClick: () => setter(e[0], e[1]),
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

const PnpReactHooksExamples: React.FunctionComponent<IPnpReactHooksExamplesProps> = (props) => {
  const [selectedExample, setExample] = React.useState<{ key: string, component: React.FunctionComponent; }>();
  const componentsLinks: INavLink[] = getComponentNavLinks((key, component) => setExample({ key, component }));
  const ExampleComponentType = selectedExample?.component;

  return (
    <PnpHookOptionProvider value={props.options}>
      <Stack horizontal verticalAlign="start" tokens={{ childrenGap: 15 }}>
        <Stack.Item>
          <Nav
            selectedKey={selectedExample?.key}
            styles={navStyles}
            groups={[{
              links: componentsLinks,
              name: "Examples",
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
    </PnpHookOptionProvider>
  );
}


export default PnpReactHooksExamples;
