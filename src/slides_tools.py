#!/usr/bin/env python3
"""
Simplified Google Slides API wrapper for AI agents.
Token-efficient interface for reading and editing presentations.
"""
# Handle both relative and absolute imports
try:
    from .drive_client import get_drive_service, get_slides_service, log_message, suppress_output
except ImportError:
    from drive_client import get_drive_service, get_slides_service, log_message, suppress_output


def get_presentation_overview(presentation_id: str) -> dict:
    """
    Get a token-efficient overview of a presentation.
    
    Args:
        presentation_id: The presentation ID
        
    Returns:
        Minimal presentation summary with slide counts
    """
    log_message(f"Getting overview for presentation: {presentation_id}")
    
    try:
        service = get_drive_service()
        creds = service._credentials
        slides_service = get_slides_service(creds)
        
        with suppress_output():
            presentation = slides_service.presentations().get(
                presentationId=presentation_id
            ).execute()
        
        # Create minimal overview
        slides_summary = []
        for idx, slide in enumerate(presentation.get('slides', []), 1):
            # Count elements and get first text as summary
            elements = slide.get('pageElements', [])
            element_count = len(elements)
            
            # Try to extract a brief summary from title or first text
            summary = f"Slide {idx}"
            for elem in elements:
                if 'shape' in elem and 'text' in elem['shape']:
                    text_elements = elem['shape']['text'].get('textElements', [])
                    for text_elem in text_elements:
                        if 'textRun' in text_elem:
                            text = text_elem['textRun'].get('content', '').strip()
                            if text:
                                summary = text[:50] + ('...' if len(text) > 50 else '')
                                break
                    break
            
            slides_summary.append({
                'num': idx,
                'summary': summary,
                'elements': element_count
            })
        
        log_message(f"Overview complete: {len(slides_summary)} slides")
        
        return {
            'id': presentation.get('presentationId'),
            'title': presentation.get('title'),
            'slide_count': len(slides_summary),
            'slides': slides_summary
        }
    
    except Exception as e:
        log_message(f"ERROR getting overview: {str(e)}")
        raise


def get_slide(presentation_id: str, slide_number: int) -> dict:
    """
    Get detailed content of a specific slide.
    
    Args:
        presentation_id: The presentation ID
        slide_number: Slide number (1-indexed)
        
    Returns:
        Slide content with simplified element structure
    """
    log_message(f"Getting slide {slide_number} from presentation: {presentation_id}")
    
    try:
        service = get_drive_service()
        creds = service._credentials
        slides_service = get_slides_service(creds)
        
        with suppress_output():
            presentation = slides_service.presentations().get(
                presentationId=presentation_id
            ).execute()
        
        slides = presentation.get('slides', [])
        if slide_number < 1 or slide_number > len(slides):
            raise ValueError(f'Slide {slide_number} not found (presentation has {len(slides)} slides)')
        
        slide = slides[slide_number - 1]
        
        # Extract elements in simplified format
        elements = []
        for page_elem in slide.get('pageElements', []):
            elem_id = page_elem.get('objectId')
            
            # Text elements (shapes with text)
            if 'shape' in page_elem:
                shape = page_elem['shape']
                if 'text' in shape:
                    # Aggregate all text runs
                    text_content = []
                    for text_elem in shape['text'].get('textElements', []):
                        if 'textRun' in text_elem:
                            text_content.append(text_elem['textRun'].get('content', ''))
                    
                    text = ''.join(text_content).strip()
                    if text:
                        # Detect role (title vs body) based on placeholder type
                        role = 'text'
                        placeholder = shape.get('placeholder', {})
                        placeholder_type = placeholder.get('type', '')
                        
                        if 'TITLE' in placeholder_type or 'CENTERED_TITLE' in placeholder_type:
                            role = 'title'
                        elif 'BODY' in placeholder_type or 'SUBTITLE' in placeholder_type:
                            role = 'body'
                        
                        elements.append({
                            'id': elem_id,
                            'type': 'text',
                            'role': role,
                            'text': text
                        })
            
            # Images
            elif 'image' in page_elem:
                image = page_elem['image']
                elements.append({
                    'id': elem_id,
                    'type': 'image',
                    'url': image.get('contentUrl', '')
                })
            
            # Tables
            elif 'table' in page_elem:
                table = page_elem['table']
                elements.append({
                    'id': elem_id,
                    'type': 'table',
                    'rows': table.get('rows', 0),
                    'cols': table.get('columns', 0)
                })
            
            # Videos
            elif 'video' in page_elem:
                video = page_elem['video']
                elements.append({
                    'id': elem_id,
                    'type': 'video',
                    'url': video.get('url', '')
                })
        
        log_message(f"Retrieved slide {slide_number} with {len(elements)} elements")
        
        return {
            'num': slide_number,
            'id': slide.get('objectId'),
            'elements': elements
        }
    
    except Exception as e:
        log_message(f"ERROR getting slide: {str(e)}")
        raise


def update_text(presentation_id: str, slide_number: int, element_id: str, text: str) -> dict:
    """
    Update text in a specific element.
    
    Args:
        presentation_id: The presentation ID
        slide_number: Slide number (1-indexed)
        element_id: Element object ID from get_slide
        text: New text content
        
    Returns:
        Success status
    """
    log_message(f"Updating text in slide {slide_number}, element {element_id}")
    
    try:
        service = get_drive_service()
        creds = service._credentials
        slides_service = get_slides_service(creds)
        
        # Build the update request
        requests = [
            {
                'deleteText': {
                    'objectId': element_id,
                    'textRange': {
                        'type': 'ALL'
                    }
                }
            },
            {
                'insertText': {
                    'objectId': element_id,
                    'text': text,
                    'insertionIndex': 0
                }
            }
        ]
        
        with suppress_output():
            slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': requests}
            ).execute()
        
        log_message(f"Successfully updated text in element {element_id}")
        return {'success': True}
    
    except Exception as e:
        log_message(f"ERROR updating text: {str(e)}")
        raise


def replace_slide_elements(presentation_id: str, slide_number: int, elements: list) -> dict:
    """
    Bulk update multiple elements on a slide.
    
    Args:
        presentation_id: The presentation ID
        slide_number: Slide number (1-indexed)
        elements: List of element updates [{"id": "...", "text": "..."}]
        
    Returns:
        Success status with count
    """
    log_message(f"Bulk updating {len(elements)} elements on slide {slide_number}")
    
    try:
        service = get_drive_service()
        creds = service._credentials
        slides_service = get_slides_service(creds)
        
        # Build batch requests for all text updates
        requests = []
        for elem in elements:
            if 'id' in elem and 'text' in elem:
                # Delete existing text
                requests.append({
                    'deleteText': {
                        'objectId': elem['id'],
                        'textRange': {'type': 'ALL'}
                    }
                })
                # Insert new text
                requests.append({
                    'insertText': {
                        'objectId': elem['id'],
                        'text': elem['text'],
                        'insertionIndex': 0
                    }
                })
        
        if requests:
            with suppress_output():
                slides_service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': requests}
                ).execute()
        
        log_message(f"Successfully updated {len(elements)} elements")
        return {'success': True, 'updated': len(elements)}
    
    except Exception as e:
        log_message(f"ERROR replacing elements: {str(e)}")
        raise


def add_element(presentation_id: str, slide_number: int, element: dict) -> dict:
    """
    Add a new element (image/table) to a slide.
    
    Args:
        presentation_id: The presentation ID
        slide_number: Slide number (1-indexed)
        element: Element spec {"type": "image", "url": "...", "position": "center"}
        
    Returns:
        Success status with new element ID
    """
    log_message(f"Adding {element.get('type')} to slide {slide_number}")
    
    try:
        service = get_drive_service()
        creds = service._credentials
        slides_service = get_slides_service(creds)
        
        # Get slide to determine page size
        with suppress_output():
            presentation = slides_service.presentations().get(
                presentationId=presentation_id
            ).execute()
        
        slides = presentation.get('slides', [])
        if slide_number < 1 or slide_number > len(slides):
            raise ValueError(f'Slide {slide_number} not found')
        
        slide = slides[slide_number - 1]
        page_id = slide.get('objectId')
        
        # Get page dimensions
        page_size = presentation.get('pageSize', {})
        page_width = page_size.get('width', {}).get('magnitude', 720)
        page_height = page_size.get('height', {}).get('magnitude', 540)
        
        # Calculate position based on simple positioning
        position = element.get('position', 'center')
        size_width = page_width * 0.6
        size_height = page_height * 0.4
        
        if position == 'top':
            translate_x = (page_width - size_width) / 2
            translate_y = page_height * 0.1
        elif position == 'bottom':
            translate_x = (page_width - size_width) / 2
            translate_y = page_height * 0.5
        else:  # center
            translate_x = (page_width - size_width) / 2
            translate_y = (page_height - size_height) / 2
        
        elem_id = f'element_{slide_number}_{element.get("type")}'
        
        requests = []
        
        if element.get('type') == 'image' and element.get('url'):
            requests.append({
                'createImage': {
                    'objectId': elem_id,
                    'url': element['url'],
                    'elementProperties': {
                        'pageObjectId': page_id,
                        'size': {
                            'width': {'magnitude': size_width, 'unit': 'PT'},
                            'height': {'magnitude': size_height, 'unit': 'PT'}
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': translate_x,
                            'translateY': translate_y,
                            'unit': 'PT'
                        }
                    }
                }
            })
        
        elif element.get('type') == 'table':
            rows = element.get('rows', 2)
            cols = element.get('cols', 2)
            requests.append({
                'createTable': {
                    'objectId': elem_id,
                    'elementProperties': {
                        'pageObjectId': page_id,
                        'size': {
                            'width': {'magnitude': size_width, 'unit': 'PT'},
                            'height': {'magnitude': size_height, 'unit': 'PT'}
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': translate_x,
                            'translateY': translate_y,
                            'unit': 'PT'
                        }
                    },
                    'rows': rows,
                    'columns': cols
                }
            })
        
        if requests:
            with suppress_output():
                slides_service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': requests}
                ).execute()
            
            log_message(f"Successfully added {element.get('type')} element")
            return {'success': True, 'element_id': elem_id}
        else:
            raise ValueError('Unsupported element type or missing parameters')
    
    except Exception as e:
        log_message(f"ERROR adding element: {str(e)}")
        raise


def duplicate_slide(presentation_id: str, source_slide: int, insert_at: int) -> dict:
    """
    Duplicate a slide.
    
    Args:
        presentation_id: The presentation ID
        source_slide: Source slide number (1-indexed)
        insert_at: Where to insert the duplicate (1-indexed)
        
    Returns:
        New slide number
    """
    log_message(f"Duplicating slide {source_slide} to position {insert_at}")
    
    try:
        service = get_drive_service()
        creds = service._credentials
        slides_service = get_slides_service(creds)
        
        with suppress_output():
            presentation = slides_service.presentations().get(
                presentationId=presentation_id
            ).execute()
        
        slides = presentation.get('slides', [])
        if source_slide < 1 or source_slide > len(slides):
            raise ValueError(f'Source slide {source_slide} not found')
        
        source_slide_id = slides[source_slide - 1].get('objectId')
        
        # Calculate insertion index (0-based)
        insertion_index = min(insert_at - 1, len(slides))
        
        requests = [{
            'duplicateObject': {
                'objectId': source_slide_id,
                'objectIds': {}
            }
        }]
        
        with suppress_output():
            response = slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': requests}
            ).execute()
        
        log_message(f"Successfully duplicated slide {source_slide}")
        
        # New slide is added after source, need to move if needed
        new_slide_number = source_slide + 1
        
        return {'success': True, 'new_slide_number': new_slide_number}
    
    except Exception as e:
        log_message(f"ERROR duplicating slide: {str(e)}")
        raise

