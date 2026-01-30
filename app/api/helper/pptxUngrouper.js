const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");

class PptxUngrouper {
  /**
   * Ungroup elements in the slide XML
   * @param {string} xmlContent - The XML content to process
   * @returns {string} - The processed XML content with ungrouped elements
   */
  static async ungroupElements(xmlContent) {
    try {
      // Clean and validate XML content before parsing
      const cleanedXml = this.cleanXmlContent(xmlContent);

      if (!cleanedXml || cleanedXml.trim().length === 0) {
        console.warn('Empty or invalid XML content provided to ungroupElements');
        return xmlContent;
      }

      // Create DOM parser and parse XML
      const parser = new DOMParser();

      // Add error handler to capture parsing issues
      const errorHandler = {
        warning: function (w) { console.warn('XML Warning:', w); },
        error: function (e) { console.error('XML Error:', e); },
        fatalError: function (e) {
          console.error('XML Fatal Error:', e);
          throw e;
        }
      };

      const doc = parser.parseFromString(cleanedXml, 'application/xml', errorHandler);

      // Check if parsing was successful
      if (!doc || doc.getElementsByTagName('parsererror').length > 0) {
        console.warn('XML parsing failed, returning original content');
        return xmlContent;
      }

      // Find the spTree (main container)
      const spTree = doc.getElementsByTagName('p:spTree')[0];
      if (!spTree) {
        return xmlContent; // No shapes to ungroup
      }

      // Collect all shapes and their absolute positions
      const shapesWithPositions = [];
      this.collectAllShapesWithPositions(doc, spTree, shapesWithPositions);

      // If we found group shapes to ungroup
      if (shapesWithPositions.length > 0) {
        console.log(`Found ${shapesWithPositions.length} shapes in groups to ungroup`);

        // Move all shapes to slide level with their absolute positions
        for (const shapeInfo of shapesWithPositions) {
          this.moveShapeToSlideLevel(shapeInfo.shape, spTree, shapeInfo.position);
        }

        // Remove all groups (they should now be empty)
        this.removeAllGroups(doc);

        // Serialize back to XML
        const serializer = new XMLSerializer();
        let updatedXml = serializer.serializeToString(doc);

        // Fix namespace issues
        updatedXml = this.fixNamespaces(updatedXml);

        return updatedXml;
      }

      return xmlContent; // No changes made
    } catch (error) {
      console.error('Error in ungroupElements:', error);
      console.warn('Returning original XML content due to parsing error');
      return xmlContent; // Return original content on error
    }
  }

  /**
   * Clean XML content before parsing to handle common issues
   * @param {string} xmlContent - Raw XML content
   * @returns {string} - Cleaned XML content
   */
  static cleanXmlContent(xmlContent) {
    if (!xmlContent || typeof xmlContent !== 'string') {
      return xmlContent;
    }

    try {
      // Remove BOM (Byte Order Mark) characters
      let cleaned = xmlContent.replace(/^\uFEFF/, '');

      // Remove any content before the first XML declaration
      const xmlDeclMatch = cleaned.match(/<\?xml[^>]*\?>/);
      if (xmlDeclMatch) {
        const xmlDeclIndex = cleaned.indexOf(xmlDeclMatch[0]);
        if (xmlDeclIndex > 0) {
          cleaned = cleaned.substring(xmlDeclIndex);
        }
      }

      // Remove duplicate XML declarations (keep only the first one)
      const xmlDeclarations = cleaned.match(/<\?xml[^>]*\?>/g);
      if (xmlDeclarations && xmlDeclarations.length > 1) {
        // Find the position of the second XML declaration and remove it
        let firstDeclEnd = cleaned.indexOf(xmlDeclarations[0]) + xmlDeclarations[0].length;
        let restOfContent = cleaned.substring(firstDeclEnd);

        // Remove any additional XML declarations
        for (let i = 1; i < xmlDeclarations.length; i++) {
          restOfContent = restOfContent.replace(xmlDeclarations[i], '');
        }

        cleaned = cleaned.substring(0, firstDeclEnd) + restOfContent;
      }

      // Trim whitespace from the beginning and end
      cleaned = cleaned.trim();

      // Basic validation - ensure it starts with XML declaration or root element
      if (!cleaned.startsWith('<?xml') && !cleaned.startsWith('<')) {
        console.warn('XML content does not start with valid XML declaration or element');
        return xmlContent; // Return original if we can't clean it properly
      }

      return cleaned;
    } catch (error) {
      console.error('Error cleaning XML content:', error);
      return xmlContent; // Return original on error
    }
  }

  /**
   * Collect all shapes with their calculated absolute positions
   * This handles any level of nesting by using a stack-based approach
   */
  static collectAllShapesWithPositions(doc, spTree, result) {
    try {
      // Find all groups
      const groups = Array.from(doc.getElementsByTagName('p:grpSp'));

      // For each group, process all shapes inside it
      for (const group of groups) {
        // Create a stack to handle nested groups and shapes
        const stack = [{
          node: group,
          transform: {
            x: 0,
            y: 0,
            scaleX: 1,
            scaleY: 1,
            rotation: 0,
            flipH: false,
            flipV: false
          },
          parentTransforms: []
        }];

        // Process all nodes in the stack
        while (stack.length > 0) {
          const current = stack.pop();
          const node = current.node;

          // Skip if not a valid node
          if (!node || !node.nodeName) continue;

          // Get node info for logging
          const nodeInfo = this.getNodeInfo(node);

          // If this is a group, get its transform and process children
          if (node.nodeName === 'p:grpSp') {
            // Get this group's transform
            const groupTransform = this.getTransform(node);

            // Combine with parent transform
            const combinedTransform = this.combineTransforms(current.transform, groupTransform);

            // Process all children of this group
            const children = node.childNodes;
            for (let i = children.length - 1; i >= 0; i--) {  // Reverse order for stack
              const child = children[i];
              if (child.nodeName && ['p:sp', 'p:cxnSp', 'p:pic', 'p:grpSp'].includes(child.nodeName)) {
                // Add all parent transforms for later calculation
                const parentTransforms = [...current.parentTransforms, groupTransform];

                // Add child to stack for processing
                stack.push({
                  node: child,
                  transform: combinedTransform,
                  parentTransforms: parentTransforms
                });
              }
            }
          }
          // If this is a shape, calculate its absolute position
          else if (['p:sp', 'p:cxnSp', 'p:pic'].includes(node.nodeName)) {
            // Get shape's own transform
            const shapeTransform = this.getTransform(node);

            // Calculate absolute position
            const position = this.calculateAbsolutePosition(shapeTransform, current.parentTransforms);

            // Add to result
            result.push({
              shape: node,
              position: position,
              info: nodeInfo
            });
          }
        }
      }
    } catch (error) {
      console.error('Error in collectAllShapesWithPositions:', error);
    }
  }

  /**
   * Get transform information from a node
   */
  static getTransform(node) {
    const transform = {
      x: 0,
      y: 0,
      width: 0,
      height: 0,
      // For groups
      chOffX: 0,
      chOffY: 0,
      chWidth: 0,
      chHeight: 0,
      // General
      rotation: 0,
      flipH: false,
      flipV: false
    };

    try {
      // Get the xfrm element
      const xfrm = node.getElementsByTagName('a:xfrm')[0];
      if (!xfrm) return transform;

      // Get position
      const off = xfrm.getElementsByTagName('a:off')[0];
      if (off) {
        transform.x = parseInt(off.getAttribute('x') || '0');
        transform.y = parseInt(off.getAttribute('y') || '0');
      }

      // Get size
      const ext = xfrm.getElementsByTagName('a:ext')[0];
      if (ext) {
        transform.width = parseInt(ext.getAttribute('cx') || '0');
        transform.height = parseInt(ext.getAttribute('cy') || '0');
      }

      // For groups, get child coordinate system
      if (node.nodeName === 'p:grpSp') {
        const chOff = xfrm.getElementsByTagName('a:chOff')[0];
        if (chOff) {
          transform.chOffX = parseInt(chOff.getAttribute('x') || '0');
          transform.chOffY = parseInt(chOff.getAttribute('y') || '0');
        }

        const chExt = xfrm.getElementsByTagName('a:chExt')[0];
        if (chExt) {
          transform.chWidth = parseInt(chExt.getAttribute('cx') || '0');
          transform.chHeight = parseInt(chExt.getAttribute('cy') || '0');
        }
      }

      // Get rotation
      if (xfrm.hasAttribute('rot')) {
        transform.rotation = parseInt(xfrm.getAttribute('rot') || '0');
      }

      // Get flip attributes
      transform.flipH = xfrm.getAttribute('flipH') === '1';
      transform.flipV = xfrm.getAttribute('flipV') === '1';
    } catch (error) {
      console.error('Error getting transform for node:', error);
    }

    return transform;
  }

  /**
   * Combine two transforms into a single transform
   */
  static combineTransforms(parentTransform, childTransform) {
    try {
      // Calculate scaling factors for the child coordinate system
      let scaleX = 1, scaleY = 1;

      if (childTransform.chWidth && childTransform.width) {
        scaleX = childTransform.chWidth !== 0 ? childTransform.width / childTransform.chWidth : 1;
      }

      if (childTransform.chHeight && childTransform.height) {
        scaleY = childTransform.chHeight !== 0 ? childTransform.height / childTransform.chHeight : 1;
      }

      // Create the combined transform
      return {
        x: parentTransform.x + childTransform.x,
        y: parentTransform.y + childTransform.y,
        scaleX: parentTransform.scaleX * scaleX,
        scaleY: parentTransform.scaleY * scaleY,
        rotation: (parentTransform.rotation + childTransform.rotation) % 21600000,
        flipH: parentTransform.flipH !== childTransform.flipH,  // XOR operation
        flipV: parentTransform.flipV !== childTransform.flipV,  // XOR operation
        // Pass through child coordinate system info
        chOffX: childTransform.chOffX,
        chOffY: childTransform.chOffY
      };
    } catch (error) {
      console.error('Error combining transforms:', error);
      return parentTransform; // Return parent transform on error
    }
  }

  /**
   * Calculate absolute position for a shape based on all parent transforms
   */
  static calculateAbsolutePosition(shapeTransform, parentTransforms) {
    try {
      // Start with the shape's own transform
      const position = {
        x: shapeTransform.x,
        y: shapeTransform.y,
        width: shapeTransform.width,
        height: shapeTransform.height,
        rotation: shapeTransform.rotation,
        flipH: shapeTransform.flipH,
        flipV: shapeTransform.flipV
      };

      // Apply each parent transform in reverse order (innermost to outermost)
      let currentX = position.x;
      let currentY = position.y;
      let currentWidth = position.width;
      let currentHeight = position.height;
      let currentRotation = position.rotation;
      let currentFlipH = position.flipH;
      let currentFlipV = position.flipV;

      // Process parent transforms in reverse (from direct parent outward)
      for (let i = parentTransforms.length - 1; i >= 0; i--) {
        const parentTransform = parentTransforms[i];

        // Adjust position based on parent's coordinate system
        const relX = currentX - parentTransform.chOffX;
        const relY = currentY - parentTransform.chOffY;

        // Calculate scaling factors
        let scaleX = 1, scaleY = 1;
        if (parentTransform.chWidth && parentTransform.width) {
          scaleX = parentTransform.chWidth !== 0 ? parentTransform.width / parentTransform.chWidth : 1;
        }
        if (parentTransform.chHeight && parentTransform.height) {
          scaleY = parentTransform.chHeight !== 0 ? parentTransform.height / parentTransform.chHeight : 1;
        }

        // Apply scaling and add parent's offset
        currentX = parentTransform.x + (relX * scaleX);
        currentY = parentTransform.y + (relY * scaleY);
        currentWidth = currentWidth * scaleX;
        currentHeight = currentHeight * scaleY;

        // Apply rotation from parent
        currentRotation = (currentRotation + parentTransform.rotation) % 21600000;

        // Apply flips from parent (XOR operation)
        currentFlipH = currentFlipH !== parentTransform.flipH;
        currentFlipV = currentFlipV !== parentTransform.flipV;
      }

      // Update the final position
      position.x = currentX;
      position.y = currentY;
      position.width = currentWidth;
      position.height = currentHeight;
      position.rotation = currentRotation;
      position.flipH = currentFlipH;
      position.flipV = currentFlipV;

      return position;
    } catch (error) {
      console.error('Error calculating absolute position:', error);
      return shapeTransform; // Return original transform on error
    }
  }

  /**
   * Move a shape to slide level with its calculated absolute position
   */
  static moveShapeToSlideLevel(shape, spTree, position) {
    try {
      if (!shape.parentNode) return;

      const shapeInfo = this.getNodeInfo(shape);

      // Clone the shape to avoid reference issues
      const clone = shape.cloneNode(true);

      // Update the shape's position and size
      this.updateShapePosition(clone, position);

      // Add to slide level
      spTree.appendChild(clone);

      // Remove original (optional, as we will remove all groups later)
      if (shape.parentNode) {
        shape.parentNode.removeChild(shape);
      }
    } catch (error) {
      console.error('Error moving shape to slide level:', error);
    }
  }

  /**
   * Update a shape's position and properties
   */
  static updateShapePosition(shape, position) {
    try {
      // Get the shape's xfrm element
      const xfrm = shape.getElementsByTagName('a:xfrm')[0];
      if (!xfrm) return;

      // Update position
      const off = xfrm.getElementsByTagName('a:off')[0];
      if (off) {
        off.setAttribute('x', Math.round(position.x).toString());
        off.setAttribute('y', Math.round(position.y).toString());
      }

      // Update size
      const ext = xfrm.getElementsByTagName('a:ext')[0];
      if (ext) {
        ext.setAttribute('cx', Math.round(position.width).toString());
        ext.setAttribute('cy', Math.round(position.height).toString());
      }

      // Update rotation
      if (position.rotation !== 0) {
        xfrm.setAttribute('rot', position.rotation.toString());
      } else {
        xfrm.removeAttribute('rot');
      }

      // Update flip attributes
      if (position.flipH) {
        xfrm.setAttribute('flipH', '1');
      } else {
        xfrm.removeAttribute('flipH');
      }

      if (position.flipV) {
        xfrm.setAttribute('flipV', '1');
      } else {
        xfrm.removeAttribute('flipV');
      }
    } catch (error) {
      console.error('Error updating shape position:', error);
    }
  }

  /**
   * Remove all group nodes from the document
   */
  static removeAllGroups(doc) {
    try {
      const groups = Array.from(doc.getElementsByTagName('p:grpSp'));

      for (const group of groups) {
        if (group.parentNode) {
          group.parentNode.removeChild(group);
        }
      }
    } catch (error) {
      console.error('Error removing groups:', error);
    }
  }

  /**
   * Get human-readable info about a node
   */
  static getNodeInfo(node) {
    if (!node) return 'unknown';

    let id = '?', name = node.nodeName;

    try {
      let cNvPrElem = null;

      // Check different parent elements based on node type
      if (node.nodeName === 'p:grpSp') {
        const nvGrpSpPr = node.getElementsByTagName('p:nvGrpSpPr')[0];
        if (nvGrpSpPr) {
          cNvPrElem = nvGrpSpPr.getElementsByTagName('p:cNvPr')[0];
        }
      } else if (node.nodeName === 'p:sp') {
        const nvSpPr = node.getElementsByTagName('p:nvSpPr')[0];
        if (nvSpPr) {
          cNvPrElem = nvSpPr.getElementsByTagName('p:cNvPr')[0];
        }
      } else if (node.nodeName === 'p:cxnSp') {
        const nvCxnSpPr = node.getElementsByTagName('p:nvCxnSpPr')[0];
        if (nvCxnSpPr) {
          cNvPrElem = nvCxnSpPr.getElementsByTagName('p:cNvPr')[0];
        }
      } else if (node.nodeName === 'p:pic') {
        const nvPicPr = node.getElementsByTagName('p:nvPicPr')[0];
        if (nvPicPr) {
          cNvPrElem = nvPicPr.getElementsByTagName('p:cNvPr')[0];
        }
      }

      if (cNvPrElem) {
        id = cNvPrElem.getAttribute('id') || '?';
        name = cNvPrElem.getAttribute('name') || node.nodeName;
      }
    } catch (e) {
      // Ignore errors in getting node info
    }

    return `${name}(id:${id})`;
  }

  /**
   * Fix XML namespace issues in the serialized XML
   */
  static fixNamespaces(xml) {
    try {
      // Fix duplicate default namespace
      xml = xml.replace(/xmlns=\"[^\"]+\"\s+xmlns=\"[^\"]+\"/g, match => match.split(/\s+xmlns=/)[0]);

      // Fix duplicate namespace for a
      xml = xml.replace(/xmlns:a=\"[^\"]+\"\s+xmlns:a=\"[^\"]+\"/g, match => match.split(/\s+xmlns:a=/)[0]);

      // Fix other namespaces
      ['p', 'r', 'a16', 'a14', 'p14', 'mc', 'xdr', 'wp'].forEach(prefix => {
        const regex = new RegExp(`xmlns:${prefix}=\\\"[^\\\"]+\\\"\\s+xmlns:${prefix}=\\\"[^\\\"]+\\\"`, 'g');
        xml = xml.replace(regex, match => match.split(new RegExp(`\\s+xmlns:${prefix}=`))[0]);
      });

      return xml;
    } catch (error) {
      console.error('Error fixing namespaces:', error);
      return xml; // Return original on error
    }
  }
}

module.exports = PptxUngrouper;