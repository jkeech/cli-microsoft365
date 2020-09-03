import commands from '../../commands';
import Command from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./completion-clink-update');
import * as assert from 'assert';
import Utils from '../../../../Utils';
import { autocomplete } from '../../../../autocomplete';

describe(commands.COMPLETION_CLINK_UPDATE, () => {
  let log: string[];
  let cmdInstance: any;
  let generateClinkCompletionStub: sinon.SinonStub;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    generateClinkCompletionStub = sinon.stub(autocomplete, 'getClinkCompletion').callsFake(() => '');
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
  });

  afterEach(() => {
    generateClinkCompletionStub.reset();
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      autocomplete.getClinkCompletion
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.COMPLETION_CLINK_UPDATE), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('builds command completion', (done) => {
    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(generateClinkCompletionStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});