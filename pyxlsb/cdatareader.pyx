from libc.stdint cimport uint8_t, uint16_t, uint32_t, int32_t
from cpython.string cimport PyString_Check, PyString_AS_STRING, PyString_GET_SIZE, PyString_FromStringAndSize
from cpython.unicode cimport PyUnicode_Decode


cdef class DataReader(object):
    cdef char* _buf
    cdef Py_ssize_t _pos
    cdef Py_ssize_t _size

    cdef char* _enc

    def __init__(self, object buf, object enc=None):
        if not PyString_Check(buf):
            raise TypeError('buf needs to be a string object')

        self._buf = PyString_AS_STRING(buf)
        self._pos = 0
        self._size = PyString_GET_SIZE(buf)

        self._enc = enc if enc is not None else 'utf-16'

    def skip(self, Py_ssize_t size):
        if size < 0:
            raise ValueError('size must be positive')

        if self._pos < self._size:
            self._pos += size
            if self._pos > self._size:
                self._pos = self._size

    def read(self, Py_ssize_t size):
        cdef object value

        if size < 0:
            raise ValueError('size must be positive')

        if size > self._size - self._pos:
            return None

        value = PyString_FromStringAndSize(&self._buf[self._pos], size)
        self._pos += size

        return value

    def read_bool(self):
        cdef uint8_t value
        if self._pos >= self._size:
            return None
        value = (<uint8_t*> &self._buf[self._pos])[0]
        self._pos += 1
        return value != 0

    def read_byte(self):
        cdef uint8_t value
        if self._pos >= self._size:
            return None
        value = (<uint8_t*> &self._buf[self._pos])[0]
        self._pos += 1
        return value

    def read_short(self):
        cdef uint16_t value
        if self._size - self._pos < 2:
            return None
        value = (<uint16_t*> &self._buf[self._pos])[0]
        self._pos += 2
        return value

    def read_int(self):
        cdef uint32_t value
        if self._size - self._pos < 4:
            return None
        value = (<uint32_t*> &self._buf[self._pos])[0]
        self._pos += 4
        return value

    def read_float(self):
        cdef float value
        if self._size - self._pos < 4:
            return None
        value = (<float*> &self._buf[self._pos])[0]
        self._pos += 4
        return value

    def read_double(self):
        cdef double value
        if self._size - self._pos < 8:
            return None
        value = (<double*> &self._buf[self._pos])[0]
        self._pos += 8
        return value

    def read_rk(self):
        cdef int32_t intval
        cdef double value

        if self._size - self._pos < 4:
            return None

        intval = (<int32_t*> &self._buf[self._pos])[0]
        self._pos += 4

        if intval & 0x02 == 0x02:
            value = <double> (intval >> 2)
        else:
            (<uint32_t*> &value)[0] = intval & 0xFFFFFFFC
            (<uint32_t*> &value)[1] = 0

        if intval & 0x01 == 0x01:
            value /= 100

        return value

    def read_string(self, object size=None, object enc=None):
        cdef object value
        cdef char* e

        if size is None:
            size = self.read_int()
            if size is None:
                return None

        if size < 0:
            raise ValueError('size must be positive')

        size *= 2

        if self._size - self._pos < size:
            return None

        if PyString_Check(enc):
            e = PyString_AS_STRING(enc)
        elif enc is not None:
            raise TypeError('enc must be a string object')
        else:
            e = self._enc

        value = PyUnicode_Decode(&self._buf[self._pos], size, e, 'strict')
        self._pos += size

        return value
