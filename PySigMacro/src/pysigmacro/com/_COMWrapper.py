#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-21 23:56:38 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_COMWrapper.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_COMWrapper.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ._inspect import inspect
from win32com.client import VARIANT
import pythoncom
from ._AccessTracker import AccessTracker
from ._base import get_wrapper

class COMWrapper:
    """
    A wrapper class that exposes COM object properties, methods and child objects
    using a pythonic interface based on inspection.
    """
    def __init__(self, com_object, path=""):
        """
        Initialize the wrapper with a COM object.
        """
        self._com_object = com_object
        self._path = path
        self._inspected = inspect(com_object)
        self._cached_objects = {}
        self._wrap_properties()
        # self._wrap_methods()  # Make sure this line exists

    # def _wrap_methods(self, name):
    #     """
    #     Create a wrapper for a COM method that properly handles return values.
    #     """
    #     def wrapped_method(*args, **kwargs):
    #         try:
    #             # Get the method from COM object
    #             com_method = getattr(self._com_object, name)
    #             # Call the method on the COM object
    #             result = com_method(*args, **kwargs)
    #             # If result is a COM object, wrap it
    #             if hasattr(result, "_oleobj_"):
    #                 # Calculate new path for method result
    #                 new_path = f"{self._path}.{name}()" if self._path else f"{name}()"
    #                 return get_wrapper(result, new_path)
    #             return result
    #         except Exception as e:
    #             print(f"Error calling method {name}: {e}")
    #             raise
    #     return wrapped_method

    def _wrap_methods(self, name):
        """
        Create a wrapper for a COM method that properly handles return values.
        """
        def wrapped_method(*args, **kwargs):
            try:
                # Get the method from COM object
                com_method = getattr(self._com_object, name)
                # Call the method on the COM object
                result = com_method(*args, **kwargs)
                # If result is a COM object, wrap it
                if hasattr(result, "_oleobj_"):
                    # Calculate new path for method result
                    new_path = f"{self._path}.{name}()" if self._path else f"{name}()"
                    return get_wrapper(result, new_path)
                return result
            except Exception as e:
                print(f"Error calling method {name}: {e}")
                raise
        return wrapped_method

    def _wrap_properties(self):
        """
        Wrap COM object properties based on inspected data.
        """
        for _, row in self._inspected.iterrows():
            if row["Type"] in ("Property", "Object"):
                self._create_property(row["Name"], row["Type"])

        # Actually make the properties accessible on the instance
        # by updating the instance's __class__ with a new class containing the properties
        cls_name = f"Dynamic{type(self).__name__}"
        bases = (type(self),)
        self.__class__ = type(cls_name, bases, {})

    def _create_property(self, name, typ):
        """
        Create a Python property wrapping the COM object's attribute.
        For 'Property' type, create a property with the given name.
        For 'Object' type, create only the property with _obj suffix.
        """
        if typ == "Property":
            def getter(instance):
                try:
                    return getattr(instance._com_object, name)
                except Exception:
                    return None

            def setter(instance, value):
                try:
                    if hasattr(value, "_com_object"):
                        value = value._com_object
                    setattr(instance._com_object, name, value)
                except Exception as e:
                    print(f"Warning: Could not set property {name}: {e}")

            setattr(type(self), name, property(getter, setter))
        elif typ == "Object":
            def getter(instance):
                try:
                    if name in instance._cached_objects:
                        return instance._cached_objects[name]

                    value = getattr(instance._com_object, name)
                    # Calculate new path for this object
                    new_path = f"{instance._path}.{name}" if instance._path else name
                    wrapped = get_wrapper(value, new_path)
                    instance._cached_objects[name] = wrapped
                    return wrapped
                except Exception:
                    return None

            def setter(instance, value):
                try:
                    if hasattr(value, "_com_object"):
                        value = value._com_object
                    setattr(instance._com_object, name, value)
                except Exception as e:
                    print(f"Warning: Could not set property {name}: {e}")

            setattr(type(self), f"{name}_obj", property(getter, setter))

    @property
    def path(self):
        return self._path

    # # Rest of your methods remain the same
    # def _wrap_methods(self):
    #     """
    #     Wrap COM object methods based on inspected data, adding them as Python methods.
    #     """
    #     for _, row in self._inspected.iterrows():
    #         if row["Type"] == "Method":
    #             name = row["Name"]
    #             method = getattr(self._com_object, name, None)
    #             if method and callable(method):
    #                 setattr(self, name, method)


    # def _wrap_methods(self):
    #     """
    #     Wrap COM object methods based on inspected data, adding them as Python methods.
    #     """
    #     for _, row in self._inspected.iterrows():
    #         if row["Type"] == "Method":
    #             name = row["Name"]

    #             def method_factory(method_name):
    #                 def wrapped_method(*args, **kwargs):
    #                     try:
    #                         # Get the method from COM object
    #                         com_method = getattr(self._com_object, method_name)
    #                         # Call the method on the COM object
    #                         result = com_method(*args, **kwargs)
    #                         # If result is a COM object, wrap it
    #                         if hasattr(result, "_oleobj_"):
    #                             # Calculate new path for method result
    #                             new_path = f"{self._path}.{method_name}()" if self._path else f"{method_name}()"
    #                             return get_wrapper(result, new_path)
    #                         return result
    #                     except Exception as e:
    #                         print(f"Error calling method {method_name}: {e}")
    #                         raise
    #                 return wrapped_method

    #             # Dynamically add method to the wrapper instance
    #             setattr(self.__class__, name, method_factory(name))


    # def __getattr__(self, name):
    #     try:
    #         attr = getattr(self._com_object, name)
    #         if callable(attr):
    #             # Wrap COM methods
    #             return self._wrap_methods(name)
    #         elif hasattr(attr, "_oleobj_"):
    #             # Wrap nested COM objects
    #             child_path = f"{self._path}.{name}"
    #             return get_wrapper(attr, child_path)
    #         else:
    #             # Return other attributes directly
    #             return attr
    #     except pythoncom.com_error as e:
    #         # This is a COM error, but the method might exist
    #         # Store it as a callable method that will pass through the error
    #         # when actually called with parameters
    #         if "member not found" in str(e).lower():
    #             # Only raise AttributeError for truly missing members
    #             raise AttributeError(f"Attribute {name} not found") from e
    #         else:
    #             # For other COM errors, create a method that will pass through the call
    #             # and raise the original error when called
    #             def method_with_error(*args, **kwargs):
    #                 raise e
    #             return method_with_error
    #     except Exception as e:
    #         raise AttributeError(f"Attribute {name} not found") from e

    # def __getattr__(self, name):
    #     try:
    #         attr = getattr(self._com_object, name)
    #         if callable(attr):
    #             # Wrap COM methods
    #             return self._wrap_methods(name)
    #         elif hasattr(attr, "_oleobj_"):
    #             # Wrap nested COM objects
    #             child_path = f"{self._path}.{name}"
    #             return get_wrapper(attr, child_path)
    #         else:
    #             # Return other attributes directly
    #             return attr
    #     except pythoncom.com_error as com_err:
    #         # This is a COM error, but the method might exist
    #         # Store it as a callable method that will pass through the error
    #         # when actually called with parameters
    #         if "member not found" in str(com_err).lower():
    #             # Only raise AttributeError for truly missing members
    #             raise AttributeError(f"Attribute {name} not found") from com_err
    #         else:
    #             # For other COM errors, create a method that will pass through the call
    #             # and raise the original error when called
    #             def method_with_error(*args, **kwargs):
    #                 # Store a reference to the error at function creation time
    #                 error = com_err
    #                 raise error
    #             return method_with_error
    #     except Exception as e:
    #         raise AttributeError(f"Attribute {name} not found") from e

    def __getattr__(self, name):
        try:
            attr = getattr(self._com_object, name)
            if callable(attr):
                # Wrap COM methods
                return self._wrap_methods(name)
            elif hasattr(attr, "_oleobj_"):
                # Wrap nested COM objects
                child_path = f"{self._path}.{name}"
                return get_wrapper(attr, child_path)
            else:
                # Return other attributes directly
                return attr
        except pythoncom.com_error as com_err:
            # This is a COM error, but the method might exist
            if "member not found" in str(com_err).lower():
                # Only raise AttributeError for truly missing members
                raise AttributeError(f"Attribute {name} not found") from com_err
            else:
                # For other COM errors, create a closure capturing the current error
                original_error = com_err  # Make a local variable for closure

                # Create a method that will raise the original error when called
                def method_with_error(*args, **kwargs):
                    raise original_error

                return method_with_error
        except Exception as e:
            raise AttributeError(f"Attribute {name} not found") from e


    def __dir__(self):
        """
        Return a list of attributes from the COM object's inspection data
        without triggering the property getters.
        """
        attrs = set()
        try:
            attrs.update(dir(self._com_object))
        except Exception:
            pass
        attrs.update(super().__dir__())
        for _, row in self._inspected.iterrows():
            typ = row["Type"]
            name = row["Name"]
            if typ == "Property":
                attrs.add(name)
            elif typ == "Object":
                attrs.add(f"{name}_obj")
            elif typ == "Method":
                attrs.add(name)
        return sorted(attrs)

    def __repr__(self):
        """
        Return a string representation of the COMWrapper.
        """
        try:
            name = getattr(self._com_object, "Name", "Unknown")
        except Exception:
            name = "Unknown"
        return f"<COMWrapper for {name} at {self._path}>"

    def __str__(self):
        """
        Return a string representation for printing.
        """
        return self.__repr__()

    def __len__(self):
        """
        Return Count for collections
        """
        if hasattr(self._com_object, "Count"):
            return self._com_object.Count
        else:
            None

    def Item(self, key):
        if key == -1 and hasattr(self._com_object, "Count"):
            key = self._com_object.Count - 1
        result = self._com_object(key)
        path = f"{self._path}({key})" if self._path else f"({key})"
        return get_wrapper(result, path)

    def __call__(self, key):
        if key == -1 and hasattr(self._com_object, "Count"):
            key = self._com_object.Count - 1
        result = self._com_object(key)
        path = f"{self._path}({key})" if self._path else f"({key})"
        return get_wrapper(result, path)

    def __getitem__(self, key):
        if key == -1 and hasattr(self._com_object, "Count"):
            key = self._com_object.Count - 1
        result = self._com_object(key)
        path = f"{self._path}[{key}]" if self._path else f"[{key}]"
        return get_wrapper(result, path)

# EOF