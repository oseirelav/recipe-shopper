const weight = ['g','kg','lb','oz'];
const volume = ['tsp','tbsp','cup','pint','quart','gal','fl oz','mL','L','stick'];
const count = ['un','unit','','clove'];
const small = ['bit','dash','pinch'];
const units = [weight,volume,count,small];

function equivalentUnits(unit1, unit2) {
  return units.some(list => list.includes(unit1.toLowerCase()) && list.includes(unit2.toLowerCase()));
}

function getWeightConversion(unit, newUnit) {
  console.log("weight conv")
  switch(unit) {
    case 'g':
      switch(newUnit) {
        case 'g':
          return 1;
        case 'kg':
          return 0.001;
        case 'lb':
          return 0.00220462;
        case 'oz':
          return 0.035274;
        default:
          console.log('error: bad unit')
          return 0;
      }
    case 'kg':
      switch(newUnit) {
        case 'g':
          return 1000;
        case 'kg':
          return 1;
        case 'lb':
          return 2.20462;
        case 'oz':
          return 35.274;
        default:
          console.log('error: bad unit')
          return 0;
      }
    case 'lb':
      switch(newUnit) {
        case 'g':
          return 453.592;
        case 'kg':
          return 0.453592;
        case 'lb':
          return 1;
        case 'oz':
          return 16;
        default:
          console.log('error: bad unit')
          return 0;
      }
    case 'oz':
      switch(newUnit) {
        case 'g':
          return 28.3495;
        case 'kg':
          return 0.0283495;
        case 'lb':
          return 1/16;
        case 'oz':
          return 1;
        default:
          console.log('error: bad unit')
          return 0;
      }
    default:
      console.log('error: bad unit');
      return 0;
  }
}

function getVolumeConversion(unit, newUnit) {
  console.log("vol conv")
  switch(unit) {
    case 'tsp':
      switch(newUnit) {
        case 'tsp':
          return 1;
        case 'tbsp':
          return 1/3;
        case 'cup':
          return 1/48;
        case 'pint':
          return 1/96;
        case 'quart':
          return 1/192;
        case 'gal':
          return 1/768;
        case 'fl oz':
          return 1/6;
        case 'mL':
          return 4.92892;
        case 'L':
          return 0.00492892;
        case 'stick':
          return 1/24;
        default:
          console.log('error: bad unit');
          return 0;
      }
    case 'tbsp':
      switch(newUnit) {
        case 'tsp':
          return 3;
        case 'tbsp':
          return 1;
        case 'cup':
          return 1/16;
        case 'pint':
          return 1/32;
        case 'quart':
          return 1/64;
        case 'gal':
          return 1/256;
        case 'fl oz':
          return 1/2;
        case 'mL':
          return 14.7868;
        case 'L':
          return 0.0147868;
        case 'stick':
          return 1/8;
        default:
          console.log('error: bad unit');
          return 0;
      }
    case 'cup':
      switch(newUnit) {
        case 'tsp':
          return 48;
        case 'tbsp':
          return 16;
        case 'cup':
          return 1;
        case 'pint':
          return 1/2;
        case 'quart':
          return 1/4;
        case 'gal':
          return 1/16;
        case 'fl oz':
          return 8;
        case 'mL':
          return 236.588;
        case 'L':
          return 0.236588;
        case 'stick':
          return 2;
        default:
          console.log('error: bad unit');
          return 0;
      }
    case 'pint':
      switch(newUnit) {
        case 'tsp':
          return 96;
        case 'tbsp':
          return 32;
        case 'cup':
          return 2;
        case 'pint':
          return 1;
        case 'quart':
          return 1/2;
        case 'gal':
          return 1/8;
        case 'fl oz':
          return 16;
        case 'mL':
          return 473.176;
        case 'L':
          return 0.473176;
        case 'stick':
          return 4;
        default:
          console.log('error: bad unit');
          return 0;
      }
    case 'quart':
      switch(newUnit) {
        case 'tsp':
          return 192;
        case 'tbsp':
          return 64;
        case 'cup':
          return 4;
        case 'pint':
          return 2;
        case 'quart':
          return 1;
        case 'gal':
          return 1/4;
        case 'fl oz':
          return 32;
        case 'mL':
          return 946.353;
        case 'L':
          return 0.946353;
        case 'stick':
          return 8;
        default:
          console.log('error: bad unit');
          return 0;
      }
    case 'gal':
      switch(newUnit) {
        case 'tsp':
          return 768;
        case 'tbsp':
          return 256;
        case 'cup':
          return 16;
        case 'pint':
          return 8;
        case 'quart':
          return 4;
        case 'gal':
          return 1;
        case 'fl oz':
          return 128;
        case 'mL':
          return 3785.41;
        case 'L':
          return 3.78541;
        case 'stick':
          return 32;
        default:
          console.log('error: bad unit');
          return 0;
      }
    case 'fl oz':
      switch(newUnit) {
        case 'tsp':
          return 6;
        case 'tbsp':
          return 2;
        case 'cup':
          return 1/8;
        case 'pint':
          return 1/16;
        case 'quart':
          return 1/32;
        case 'gal':
          return 1/128;
        case 'fl oz':
          return 1;
        case 'mL':
          return 29.5735;
        case 'L':
          return 0.0295735;
        case 'stick':
          return 1/4;
        default:
          console.log('error: bad unit');
          return 0;
      }
    case 'mL':
      switch(newUnit) {
        case 'tsp':
          return 0.202884;
        case 'tbsp':
          return 0.067628;
        case 'cup':
          return 0.00422675;
        case 'pint':
          return 0.00211338;
        case 'quart':
          return 0.00105669;
        case 'gal':
          return 0.000264172;
        case 'fl oz':
          return 0.033814;
        case 'mL':
          return 1;
        case 'L':
          return 0.001;
        case 'stick':
          return 0.00845351;
        default:
          console.log('error: bad unit');
          return 0;
      }
    case 'L':
      switch(newUnit) {
        case 'tsp':
          return 202.884;
        case 'tbsp':
          return 67.628;
        case 'cup':
          return 4.22675;
        case 'pint':
          return 2.11338;
        case 'quart':
          return 1.05669;
        case 'gal':
          return 0.264172;
        case 'fl oz':
          return 33.814;
        case 'mL':
          return 1000;
        case 'L':
          return 1;
        case 'stick':
          return 8.45351;
        default:
          console.log('error: bad unit');
          return 0;
      }
    case 'stick':
      switch(newUnit) {
        case 'tsp':
          return 24;
        case 'tbsp':
          return 8;
        case 'cup':
          return 1/2;
        case 'pint':
          return 1/4;
        case 'quart':
          return 1/8;
        case 'gal':
          return 1/32;
        case 'fl oz':
          return 4;
        case 'mL':
          return 118.294;
        case 'L':
          return 0.118294;
        case 'stick':
          return 1;
        default:
          console.log('error: bad unit');
          return 0;
      }
    default:
      console.log('error: bad unit');
      return 0;
  }
}

function getConversionRateHelper(unit, newUnit, arr) {
  switch(true) {
    case arr === count || arr === small:
      return 1;
    case arr === weight:
      return getWeightConversion(unit, newUnit);
    case arr === volume:
      return getVolumeConversion(unit, newUnit);
    default:
      console.log('error: not a valid array');
      return 0;
  }
}

function getConversionRate(unit, newUnit) {
  if (!equivalentUnits(unit,newUnit)) {
    console.log('error: units are not equivalent');
    return 0;
  }
  if (unit == newUnit) {
    console.log('same unit');
    return 1;
  }  

  for (const arr of units) {
    if (arr.includes(unit)) {
      return getConversionRateHelper(unit, newUnit, arr);
    }
  }
  console.log('error: should not get here at all')
  return 0;
}

function convert(number, conversionRate) {
  `convert to different units`
  return Math.round(number*conversionRate*100)/100;
}

function getNewWeightUnit(number, unit) {
  `choose weight conversion`
  switch(unit) {
    case 'g':
      if (number >= 453.592) {
        return 'lb';
      }
      else {
        return 'oz';
      }
    case 'kg':
      if (number >= 0.453592) {
        return 'lb';
      }
      else {
        return 'oz';
      }
    case 'lb':
      if (number >= 1) {
        return 'lb';
      }
      else {
        return 'oz';
      }
    case 'oz':
      if (number >= 16) {
        return 'lb';
      }
      else {
        return 'oz';
      }
    default:
      console.log('not a valid unit');
      return "error";
  }
}

function getNewVolumeUnit(number, unit, item) {
  `choose volume conversion`
  if (String(item).toLowerCase().endsWith("butter")) {
    return 'stick';
  }
  if (unit == 'mL') {
    if (number >= 1000) {
      return 'L';
    }
    return unit;
  }
  else if (unit == 'L') {
    if (number < 1) {
      return 'mL';
    }
    return unit;
  }
  else {
    return 'fl oz';
  }
}

function getNewUnit(number, unit, item) {
  `choose conversion`
  newUnit = ""
  switch(true) {
    case count.includes(unit) || small.includes(unit):
      return [number,unit];
    case weight.includes(unit):
      newUnit = getNewWeightUnit(number, unit);
      break;
    case volume.includes(unit):
      newUnit = getNewVolumeUnit(number, unit, item);
      break;
    default:
      console.log('not a valid unit');
      break;
  }
  const newNum = convert(number, getConversionRate(unit, newUnit));
  return [newNum,newUnit];
}

function isUnit(unit) {
  return units.some(list => list.includes(unit));
}

