const fs = require('fs');
const os = require('os');
const path = require('path');
const { Readable, Writable } = require('stream');
const AdmZip = require('adm-zip');
const { Workbook } = require('../wrapper');

function createWorkbook() {
  const workbook = new Workbook();
  const worksheet = workbook.addWorksheet('Data');
  worksheet.write(0, 0, 'streamed value');
  return workbook;
}

function createSink() {
  return new Writable({
    write(_chunk, _encoding, callback) {
      callback();
    },
  });
}

describe('Workbook.saveToStream', () => {
  let temporaryDirectory;

  beforeEach(() => {
    temporaryDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'rusc-xlsx-stream-test-'));
    jest.spyOn(os, 'tmpdir').mockReturnValue(temporaryDirectory);
  });

  afterEach(() => {
    jest.restoreAllMocks();
    fs.rmSync(temporaryDirectory, { recursive: true, force: true });
  });

  function expectTemporaryDirectoryEmpty() {
    expect(fs.readdirSync(temporaryDirectory)).toEqual([]);
  }

  test('returns a promise that resolves after the writable finishes', async () => {
    const chunks = [];
    const writable = new Writable({
      write(chunk, _encoding, callback) {
        chunks.push(Buffer.from(chunk));
        setImmediate(callback);
      },
    });

    const completion = createWorkbook().saveToStream(writable);

    expect(completion).toBeInstanceOf(Promise);
    expect(writable.writableFinished).toBe(false);

    await completion;

    expect(writable.writableFinished).toBe(true);
    const archive = new AdmZip(Buffer.concat(chunks));
    expect(archive.readAsText('xl/sharedStrings.xml')).toContain('streamed value');
    expectTemporaryDirectoryEmpty();
  });

  test('propagates native save errors and removes partial files', async () => {
    const saveError = new Error('native save failed');
    const workbook = Object.create(Workbook.prototype);
    workbook._native = {
      save(filename) {
        fs.writeFileSync(filename, 'partial workbook');
        throw saveError;
      },
    };

    const completion = workbook.saveToStream(createSink());

    expect(completion).toBeInstanceOf(Promise);
    await expect(completion).rejects.toBe(saveError);
    expectTemporaryDirectoryEmpty();
  });

  test('propagates read stream errors and removes the temporary file', async () => {
    const readError = new Error('read stream failed');
    const workbook = Object.create(Workbook.prototype);
    workbook._native = {
      save(filename) {
        fs.writeFileSync(filename, 'workbook');
      },
    };
    jest.spyOn(fs, 'createReadStream').mockReturnValue(new Readable({
      read() {
        this.destroy(readError);
      },
    }));

    await expect(workbook.saveToStream(createSink())).rejects.toBe(readError);
    expectTemporaryDirectoryEmpty();
  });

  test('propagates writable stream errors and removes the temporary file', async () => {
    const writableError = new Error('writable stream failed');
    const writable = new Writable({
      write(_chunk, _encoding, callback) {
        callback(writableError);
      },
    });

    await expect(createWorkbook().saveToStream(writable)).rejects.toBe(writableError);
    expectTemporaryDirectoryEmpty();
  });
});
