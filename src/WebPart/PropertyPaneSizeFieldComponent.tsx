import * as React from 'react';
import { Dropdown, IDropdownOption, Stack, TextField, Text } from '@fluentui/react';

const makeSplitValue = (value: string) => {
  const matches = value?.match(/(\d+)\s*(\w+|%)?/);
  return {
    number: matches?.[1] ?? '',
    units: matches?.[2] ?? ''
  };
};

export function PropertyPaneSizeFieldComponent(props: {
  value: string;
  setValue: (value: string) => void;
  getDefaultValue: () => Promise<string>;
  label: string;
  description: string;
  screenUnits: string;
}) {

  const screen = `v${props.screenUnits}`;

  const unitsOptions: IDropdownOption[] = [
    { key: screen, text: "% of the screen" },
    { key: '%', text: "% of the frame" },
    { key: 'cm', text: "centimeters" },
    { key: 'in', text: "inches" },
    { key: 'mm', text: "millimeters" },
    { key: 'pt', text: "points" },
    { key: 'px', text: "pixels" },
  ];

  const splitValue = makeSplitValue(props.value);
  const [number, setNumber] = React.useState(splitValue.number);
  const [units, setUnits] = React.useState(splitValue.units);

  const debounce = React.useRef(0);

  const [valueToSave, setValueToSave] = React.useState(props.value);
  React.useEffect(() => {
    if (valueToSave) {
      props.setValue(valueToSave);
    }
  }, [valueToSave]);

  const processChanges = () => {
    if (number && units) {
      setValueToSave(number + units);
    } else {
      props.getDefaultValue().then(defaultValue => {
        const splitDefaultValue = makeSplitValue(defaultValue);
        setNumber(splitDefaultValue.number);
        setUnits(splitDefaultValue.units);
        setValueToSave(defaultValue);
      });
    }
    debounce.current = 1000;
  };

  React.useEffect(() => {
    const timeout = setTimeout(processChanges, debounce.current);
    return () => clearTimeout(timeout);
  }, [number, units]);

  const onNumberChanged = async (_, val) => {
    setNumber(val);
  };

  const onUnitChanged = (_, val) => {
    setUnits(val.key);
  };

  return (
    <Stack tokens={{ childrenGap: "s2" }}>
      <Stack horizontal tokens={{ childrenGap: "s2" }}>
        <Stack.Item grow>
          <TextField label={props.label} value={number} onChange={onNumberChanged} />
        </Stack.Item>
        <Stack.Item align='end'>
          <Dropdown style={{ minWidth: "10em" }} options={unitsOptions} selectedKey={units} disabled={number === ''} onChange={onUnitChanged} />
        </Stack.Item>
      </Stack>
      <Text variant='small' >{props.description}</Text>
    </Stack>
  );
}
