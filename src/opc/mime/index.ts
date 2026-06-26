import standardTypes from './types';
import Mime from './Mime';

export { default as Mime } from './Mime';

export default new Mime(standardTypes)._freeze();
