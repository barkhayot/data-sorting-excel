from import_export import resources
from . models import ProductData

class ProductResource(resources.ModelResource):
    class Meta:
        model = ProductData