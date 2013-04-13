from DataTypes import DataTypes
from nose.tools import eq_

def test__to_enumeration_value():
    eq_(DataTypes.to_enumeration_value(DataTypes.BOOLEAN), "b")
    eq_(DataTypes.to_enumeration_value(DataTypes.DATE), "d")
    eq_(DataTypes.to_enumeration_value(DataTypes.ERROR), "e")
    eq_(DataTypes.to_enumeration_value(DataTypes.INLINE_STRING), "inlineStr")
    eq_(DataTypes.to_enumeration_value(DataTypes.NUMBER), "n")
    eq_(DataTypes.to_enumeration_value(DataTypes.SHARED_STRING), "s")
    eq_(DataTypes.to_enumeration_value(DataTypes.STRING), "str")

def test__get_type():
    eq_(DataTypes.get_type(15), DataTypes.NUMBER)
    eq_(DataTypes.get_type(15.0), DataTypes.NUMBER)
    eq_(DataTypes.get_type("test"), DataTypes.INLINE_STRING)

